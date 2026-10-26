from playwright.sync_api import sync_playwright, expect
from pathlib import Path
import re
import config
import log_manager as lm
from invoice_rules import is_email_pdf_spv

EDGE_PATH      = config.NIS_EDGE_PATH
LOGIN_URL      = config.NIS_LOGIN_URL
AUTH_FILE      = config.NIS_AUTH_FILE   # kept for reference / legacy
NIS_USERNAME   = config.NIS_USERNAME
NIS_PASSWORD   = config.NIS_PASSWORD
NIS_OTP_SECRET = config.NIS_OTP_SECRET
NIS_APP_IFRAME_TIMEOUT_MS = getattr(config, "NIS_APP_IFRAME_TIMEOUT_MS", 90000)


# ─────────────────────────────────────────────────────────────────────────────
# Password-based login  (replaces cookie / storage_state approach)
# ─────────────────────────────────────────────────────────────────────────────
def _login_with_password(page) -> None:
    """
    Perform the full NIS portal login flow using username + password + RSA OTP.

    Mirrors the standalone loginwithpassword.py script so the same logic
    runs inside book_nis() without needing a pre-saved cookies.json file.

    Steps
    -----
    1. Navigate to the NIS Launchpad URL.
    2. Enter email and click Next.
    3. Enter password and click Sign in.
    4. Click the 'ExternalAuth' (Approve with RSAEntraMFA) card.
    5. Fill the OTP and click Verify.
    6. Click the final 'Yes / Stay signed in' button.
    """
    print("[Login] Navigating to NIS portal …")
    page.goto(LOGIN_URL, wait_until="networkidle")

    # ── Step 1: email ────────────────────────────────────────────────────────
    page.wait_for_selector("#i0116", state="visible", timeout=30000)
    page.fill("#i0116", NIS_USERNAME)
    page.click("#idSIButton9")
    print(f"[Login] Email entered: {NIS_USERNAME}")

    # ── Step 2: password ─────────────────────────────────────────────────────
    page.wait_for_selector("#i0118", state="visible", timeout=20000)
    page.fill("#i0118", NIS_PASSWORD)
    page.click("#idSIButton9")
    page.wait_for_load_state("networkidle")
    print("[Login] Password entered.")

    # ── Step 3: ExternalAuth / RSAEntraMFA card ───────────────────────────────
    try:
        page.wait_for_selector("div[data-value='ExternalAuth']", state="visible", timeout=20000)
        page.click("div[data-value='ExternalAuth']")
        print("[Login] ExternalAuth card clicked.")
    except Exception as e:
        print(f"[Login] ExternalAuth card not found (may have been skipped): {e}")

    # ── Step 4: RSA OTP ──────────────────────────────────────────────────────
    # The SecurID form enables the Verify button only after its own JS sees the
    # OTP field change. page.fill() alone frequently does NOT trigger that
    # validation, so the button stays aria-disabled and the old click() was a
    # no-op — the OTP was never submitted and login failed intermittently.
    # We therefore: type with real keystrokes, fire input/change/keyup, poll
    # until the button is truly enabled, fall back to Enter, and confirm the
    # screen advanced — retrying the whole entry a few times.
    otp_submitted = False
    for otp_attempt in range(1, 4):
        try:
            page.wait_for_selector("#input_otp_secret", state="visible", timeout=20000)
            otp_field = page.locator("#input_otp_secret")

            # Enter the OTP with genuine key events so the form's validation runs.
            otp_field.click()
            otp_field.fill("")
            otp_field.type(str(NIS_OTP_SECRET), delay=120)
            for evt in ("input", "change", "keyup", "blur"):
                try:
                    otp_field.dispatch_event(evt)
                except Exception:
                    pass

            # Poll (~10s) for the Verify button to become genuinely enabled.
            verify = page.locator("#btn_verify_securid")
            enabled = False
            for _ in range(33):
                aria = (verify.get_attribute("aria-disabled") or "").strip().lower()
                try:
                    disabled_prop = verify.is_disabled()
                except Exception:
                    disabled_prop = False
                if aria != "true" and not disabled_prop:
                    enabled = True
                    break
                page.wait_for_timeout(300)

            if enabled:
                verify.click()
                print(f"[Login] OTP submitted via Verify button (attempt {otp_attempt}).")
            else:
                # Fallback: most SecurID forms also submit on Enter.
                print("[Login] Verify button stayed disabled — submitting OTP with Enter.")
                otp_field.press("Enter")

            # Confirm the OTP screen actually advanced; otherwise retry.
            try:
                page.wait_for_selector("#input_otp_secret", state="detached", timeout=8000)
                otp_submitted = True
            except Exception:
                page.wait_for_load_state("networkidle")
                field = page.locator("#input_otp_secret")
                if field.count() == 0 or not field.first.is_visible():
                    otp_submitted = True

            if otp_submitted:
                print(f"[Login] OTP entered and accepted (attempt {otp_attempt}).")
                break
            print(f"[Login] OTP not accepted on attempt {otp_attempt} — retrying …")
        except Exception as e:
            print(f"[Login] OTP attempt {otp_attempt} error: {e}")
        page.wait_for_timeout(1000)

    if not otp_submitted:
        print("[Login] ⚠️ OTP submission could not be confirmed after 3 attempts.")
    try:
        page.wait_for_load_state("networkidle")
    except Exception:
        pass

    # ── Step 5: 'Yes / Stay signed in' ───────────────────────────────────────
    try:
        page.wait_for_selector("#idSIButton9", state="visible", timeout=20000)
        with page.expect_navigation(wait_until="networkidle"):
            page.click("#idSIButton9")
        print("[Login] 'Yes' / Stay signed-in button clicked.")
    except Exception as e:
        print(f"[Login] Final Yes button step skipped or failed: {e}")

    print("[Login] Login flow complete — portal should now be loaded.")

NIS_MANAGE_URL = (
    "https://ael-nis-prd.launchpad.cfapps.in30.hana.ondemand.com/site"
    "?siteId=6b54bd07-1384-49b1-a4e9-781f4f23eef1"
    "&sap-ushell-config=lean"
    "#Nis-manage?sap-ui-app-id-hint=saas_approuter_com.adani.nis.managenis"
)


def apply_viewport_fix(pg):
    """Force the SAP portal to a fixed 500x500 layout so automation clicks land correctly."""
    pg.evaluate("""
    () => {
        // RESET PAGE
        document.documentElement.style.margin = "0";
        document.documentElement.style.padding = "0";
        document.documentElement.style.width = "100vw";
        document.documentElement.style.height = "100vh";
        document.documentElement.style.overflow = "hidden";

        document.body.style.margin = "0";
        document.body.style.padding = "0";
        document.body.style.width = "100vw";
        document.body.style.height = "100vh";
        document.body.style.overflow = "hidden";
        document.body.style.background = "#eaeaea";

        // FIND SAP ROOT CONTAINER
        const app =
            document.querySelector("#shell-container") ||
            document.querySelector(".sapUshellShell") ||
            document.body.firstElementChild;

        if (!app) return;

        // FIXED SIZE WEBSITE
        app.style.position = "absolute";
        app.style.top = "0";
        app.style.left = "0";
        app.style.width = "500px";
        app.style.height = "500px";
        app.style.maxWidth = "500px";
        app.style.maxHeight = "500px";
        app.style.overflow = "hidden";
        app.style.transform = "none";

        // FORCE SAP TO RECALC LAYOUT
        window.dispatchEvent(new Event("resize"));
    }
    """)


def open_view_create(page):
    """Navigate directly to the NIS Manage app URL (bypasses tile click)."""

    print("page : ", page)
    page.goto(NIS_MANAGE_URL, wait_until="networkidle")
    page.wait_for_load_state("networkidle")
    apply_viewport_fix(page)
    print("Navigated directly to NIS Manage app URL.")


def click_create_nis(page):
    """
    Inside the embedded Fiori app iframe, click the 'Create NIS' button.
    Uses role/name first (stable), then text, then a CSS id fallback.
    """
    print("1 : clicking on the nis booking")
    iframe_selector = (
        "iframe[id^='application-Nis-manage-'], "
        "iframe[name^='application-Nis-manage-'], "
        "iframe[src*='Nis-manage']"
    )
    print("2 : clicking on the nis booking")
    try:
        page.wait_for_selector(
            iframe_selector,
            state="visible",
            timeout=NIS_APP_IFRAME_TIMEOUT_MS,
        )
    except Exception as first_exc:
        lm.log_warn(
            f"NIS_APP_IFRAME_TIMEOUT timeout_ms={NIS_APP_IFRAME_TIMEOUT_MS}; "
            "reopening View & Create once"
        )
        print(
            "[NIS] Application did not load in time. "
            "Navigating directly to NIS Manage URL once..."
        )
        try:
            page.goto(NIS_MANAGE_URL, wait_until="networkidle", timeout=60000)
            page.wait_for_load_state("networkidle")
            page.wait_for_selector(
                iframe_selector,
                state="visible",
                timeout=NIS_APP_IFRAME_TIMEOUT_MS,
            )
            lm.log("NIS_APP_IFRAME_RECOVERED")
        except Exception as recovery_exc:
            raise RuntimeError(
                "NIS application iframe did not load after reopening View & Create; "
                "the portal may be slow, unavailable, or the login session may have expired. "
                f"Initial error: {first_exc}; recovery error: {recovery_exc}"
            ) from recovery_exc
    app_frame = page.frame_locator(iframe_selector)

    create_btn = app_frame.get_by_role("button", name=re.compile(r"^\s*Create\s+NIS\s*$", re.I))
    print("3 : button click : ",create_btn)
    if create_btn.count() == 0:
        print("4 : clicking on the nis booking")
        create_btn = app_frame.locator("button:has-text('Create NIS'), [role='button']:has-text('Create NIS')")

    if create_btn.count() == 0:
        print("6 : clicking on the nis booking")
        create_btn = app_frame.locator("#application-Nis-manage-component---DisplayNIS--createButton")

    create_btn.first.click()
    print("✅ 'Create NIS' clicked.")


def close_failed_employee_dialog(page, timeout=30000):
    """
    Closes the 'Failed to load Employee details' message dialog inside the NIS app iframe.
    """
    iframe_selector = (
        "iframe[id^='application-Nis-manage-'], "
        "iframe[name^='application-Nis-manage-'], "
        "iframe[src*='Nis-manage']"
    )
    page.wait_for_selector(iframe_selector, state="visible", timeout=timeout)
    frame = page.frame_locator(iframe_selector)

    error_text_pattern = re.compile(r"Failed\s+to\s+load\s+Employee\s+details", re.I)

    dialog_container = frame.locator(
        "[role='dialog'], [role='alertdialog'], .sapMDialog, .sapMMessageBox"
    ).filter(has_text=error_text_pattern)

    expect(dialog_container.first).to_be_visible(timeout=timeout)

    close_btn = dialog_container.get_by_role("button", name=re.compile(r"^\s*(Close|OK)\s*$", re.I))
    if close_btn.count() == 0:
        close_btn = dialog_container.locator(
            "#__mbox-btn-0, #__mbox-btn-1, button:has-text('Close'), button:has-text('OK'), [role='button']:has-text('Close'), [role='button']:has-text('OK')"
        )

    expect(close_btn.first).to_be_visible(timeout=timeout)
    try:
        expect(close_btn.first).to_be_enabled(timeout=timeout)
    except:
        pass

    close_btn.first.click()

    expect(dialog_container.first).not_to_be_visible(timeout=timeout)
    print("✅ Error dialog closed: 'Failed to load Employee details'.")


def fill_sub_doc_type(page, frame, value_text: str = "Non PO Vendor Payment", timeout=20000):
    """
    Select 'Sub Document Type' from the UI5 ComboBox.
    """
    combo_input = frame.locator(
        "#application-Nis-manage-component---CreateNIS--subDocType-inner"
    )
    if combo_input.count() == 0:
        combo_input = frame.get_by_role("combobox").filter(
            has=frame.get_by_label(re.compile(r"Sub\s*Document\s*Type", re.I))
        )
        if combo_input.count() == 0:
            combo_input = frame.get_by_role("combobox").first

    expect(combo_input.first).to_be_visible(timeout=timeout)
    selected = False

    arrow_btn = frame.locator(
        "#application-Nis-manage-component---CreateNIS--subDocType-arrow"
    )
    if arrow_btn.count() > 0:
        try:
            arrow_btn.click()
            page.wait_for_timeout(1000)

            popup = frame.locator(
                "#application-Nis-manage-component---CreateNIS--subDocType-popup, "
                "[role='dialog'], .sapMDialog, .sapMPopover, .sapMComboBoxBasePicker, "
                "[role='listbox']"
            )
            if popup.count() > 0:
                try:
                    expect(popup.first).to_be_visible(timeout=5000)
                except Exception:
                    pass

                search_input = popup.locator(
                    "input[type='search'], input[type='text'], "
                    ".sapMSearchFieldInner input, .sapMSFI input"
                ).first
                if search_input.count() > 0:
                    try:
                        search_input.click()
                        search_input.fill("")
                        search_input.type("Non PO", delay=30)
                        page.wait_for_timeout(500)
                    except Exception:
                        pass

                item = popup.locator(
                    f"[role='option']:has-text('{value_text}'), "
                    f"[role='listitem']:has-text('Non PO'), "
                    f"li:has-text('{value_text}'), "
                    f".sapMLIB:has-text('Non PO'), "
                    f".sapMComboBoxItem:has-text('Non PO')"
                )
                if item.count() == 0:
                    item = popup.get_by_text(re.compile(r"Non\s+PO\s+Vendor", re.I))

                if item.count() > 0:
                    try:
                        item.first.click()
                        selected = True
                        print(f"✅ Sub Document Type selected from dialog picker: {value_text}")
                    except Exception:
                        pass

                if not selected:
                    try:
                        page.keyboard.press("Escape")
                        page.wait_for_timeout(300)
                    except Exception:
                        pass
        except Exception as e:
            print(f"[SubDocType] Dialog picker approach failed: {e}")

    if not selected:
        try:
            combo_input.first.click()
            combo_input.first.fill("")
            combo_input.first.type("Non PO", delay=50)
            page.wait_for_timeout(800)

            listbox_item = frame.locator(
                "[role='option']:has-text('Non PO'), "
                ".sapMComboBoxItem:has-text('Non PO'), "
                ".sapMLIB:has-text('Non PO'), "
                "li:has-text('Non PO Vendor')"
            )
            if listbox_item.count() > 0:
                listbox_item.first.click()
                selected = True
                print(f"✅ Sub Document Type selected from type-ahead: {value_text}")
            else:
                combo_input.first.fill("")
                combo_input.first.type(value_text, delay=15)
                combo_input.first.press("Enter")
                page.wait_for_timeout(300)
                selected = True
                print(f"✅ Sub Document Type typed and committed: {value_text}")
        except Exception as e:
            print(f"[SubDocType] Type-ahead approach failed: {e}")

    if not selected:
        print("[SubDocType] All UI approaches failed, forcing value via JS...")
        combo_input.first.evaluate("""
            (el, val) => {
                var nativeInputValueSetter = Object.getOwnPropertyDescriptor(
                    window.HTMLInputElement.prototype, 'value'
                ).set;
                nativeInputValueSetter.call(el, val);
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
                var sId = el.id.replace('-inner', '');
                if (typeof sap !== 'undefined' && sap.ui && sap.ui.getCore) {
                    var oControl = sap.ui.getCore().byId(sId);
                    if (oControl) {
                        var items = oControl.getItems ? oControl.getItems() : [];
                        for (var i = 0; i < items.length; i++) {
                            var txt = items[i].getText ? items[i].getText() : '';
                            if (txt.toLowerCase().indexOf('non po') >= 0) {
                                oControl.setSelectedItem(items[i]);
                                if (oControl.fireChange) oControl.fireChange({ value: txt });
                                if (oControl.fireSelectionChange) {
                                    oControl.fireSelectionChange({ selectedItem: items[i] });
                                }
                                break;
                            }
                        }
                        if (oControl.getValue && oControl.getValue() !== val) {
                            oControl.setValue(val);
                            if (oControl.fireChange) oControl.fireChange({ value: val });
                        }
                    }
                }
            }
        """, value_text)
        print(f"✅ Sub Document Type forced via JS: {value_text}")

    try:
        expect(combo_input.first).to_have_value(
            re.compile(re.escape(value_text), re.I), timeout=timeout
        )
    except Exception:
        expect(combo_input.first).to_have_value(
            re.compile(r"Non\s*PO", re.I), timeout=timeout
        )

    print(f"✅ Sub Document Type confirmed: {value_text}")


def fill_company_code(page, frame, company_code: str = "6060", timeout: int = 20000):
    cc_input = frame.locator("#application-Nis-manage-component---CreateNIS--companyCodeInput-inner")

    if cc_input.count() == 0:
        cc_input = frame.get_by_label(re.compile(r"^\s*Company\s*Code\s*$", re.I))
    if cc_input.count() == 0:
        cc_input = frame.locator("input[placeholder='Select Company Code']").first

    expect(cc_input.first).to_be_visible(timeout=timeout)
    cc_input.first.click()
    try:
        page.wait_for_timeout(100)
    except Exception:
        pass

    cc_input.first.fill("")
    cc_input.first.type(str(company_code), delay=30)

    popup_list = frame.locator("[role='listbox'], .sapMComboBoxBasePicker, .sapMPopup-CTX, .sapMDialog")
    try:
        expect(popup_list.first).to_be_visible(timeout=800)
        wanted = frame.get_by_role("option", name=re.compile(rf"^\s*{re.escape(str(company_code))}\s*$"))
        if wanted.count() == 0:
            wanted = popup_list.locator(
                f".sapMComboBoxItem:has-text('{company_code}'), "
                f".sapMLIB:has-text('{company_code}'), "
                f"tr:has(td:has-text('{company_code}'))"
            )
        if wanted.count() > 0:
            wanted.first.click()
        else:
            cc_input.first.press("Enter")
    except Exception:
        try:
            cc_input.first.press("Enter")
        except Exception:
            pass
        cc_input.first.press("Tab")

    expect(cc_input.first).to_have_value(str(company_code), timeout=timeout)
    print(f"✅ Company Code filled: {company_code}")


def get_nis_frame(page, timeout=NIS_APP_IFRAME_TIMEOUT_MS):
    """Return a frame locator pointing to the embedded NIS app iframe."""
    iframe_selector = (
        "iframe[id^='application-Nis-manage-'], "
        "iframe[name^='application-Nis-manage-'], "
        "iframe[src*='Nis-manage']"
    )
    page.wait_for_selector(iframe_selector, state="visible", timeout=timeout)
    return page.frame_locator(iframe_selector)


def disable_ai_document_read_switch(frame, timeout: int = 10000):
    """
    Turn OFF the "Enable to automatically extract invoice data using AI"
    (aiDocumentReadSwitch) toggle before uploading the header PDF.

    The switch is a SAPUI5 sap.m.Switch rendered with role="switch" and
    aria-checked reflecting its state ("true" = On/Yes, "false" = Off/No).
    We only click it if it is currently ON so we don't accidentally re-enable
    a switch that's already off.
    """
    switch = frame.locator(
        "#application-Nis-manage-component---CreateNIS--aiDocumentReadSwitch"
    )
    try:
        switch.wait_for(state="attached", timeout=timeout)
    except Exception:
        print("✅ AI document-read switch not found, skipping disable step.")
        return

    try:
        checked = (switch.get_attribute("aria-checked") or "").strip().lower()
    except Exception:
        checked = ""

    if checked == "false":
        print("✅ AI document-read switch already OFF.")
        return

    try:
        switch.click()
        # Confirm it flipped to OFF.
        try:
            expect(switch).to_have_attribute("aria-checked", "false", timeout=timeout)
        except Exception:
            pass
        print("✅ AI document-read switch disabled (set to No).")
    except Exception as exc:
        print(f"⚠️ Could not disable AI document-read switch: {exc}")


def upload_header_pdf(page, frame, file_path: str, timeout: int = 30000, click_upload_button: bool = True, wait_for_post: bool = True):
    """
    Upload header attachment PDF using the SAPUI5 FileUploader.
    """
    if not file_path:
        print("✅ No PDF provided, skipping header upload.")
        return
        
    p = Path(file_path)
    if not p.exists():
        print(f"✅ Attachment not found ({p}), skipping header upload.")
        return
    if p.suffix.lower() != ".pdf":
        print(f"✅ Expected a PDF file, got: {p.suffix}, skipping header upload.")
        return

    # Disable the AI auto-extract toggle BEFORE uploading the PDF.
    disable_ai_document_read_switch(frame, timeout=timeout)

    # The SAPUI5 FileUploader renders the real <input type="file"> inside a
    # "sapUiFupInputMask" div which is visually hidden (overlaid behind the
    # button). So we must NOT wait for visibility — Playwright can set files on
    # a hidden input. We only wait for it to be ATTACHED to the DOM.
    form = frame.locator("#application-Nis-manage-component---CreateNIS--CreateHdrFileUploader-fu_form")
    if form.count() == 0:
        form = frame.locator("form[id$='-fu_form']").filter(
            has=frame.locator("input[type='file'][id$='-fu']")
        ).first

    file_input = frame.locator("#application-Nis-manage-component---CreateNIS--CreateHdrFileUploader-fu")
    if file_input.count() == 0:
        file_input = form.locator("input[type='file'][id$='-fu']").first

    # Wait for the input to exist in the DOM (attached), not to be visible.
    file_input.wait_for(state="attached", timeout=timeout)

    # ── Retry loop ──────────────────────────────────────────────────────────
    # SAPUI5 FileUploader internally replaces (clones) the <input type="file">
    # element when it handles the 'change' event.  Playwright's set_input_files()
    # dispatches 'change' natively, which can trigger SAPUI5's handler and swap
    # out the DOM node — making the old element's .files appear empty.
    #
    # Strategy:
    #   1. Do NOT manually dispatch a second 'change' event (Playwright does it).
    #   2. After set_input_files(), re-query the file input from the DOM so we
    #      are always inspecting the *current* node, not a detached one.
    #   3. Use the HdrAttachmentInput text field as the authoritative signal
    #      that SAPUI5 accepted the file.
    #   4. Retry up to 3 times, re-locating the input each time.
    # ────────────────────────────────────────────────────────────────────────

    MAX_UPLOAD_RETRIES = 3
    upload_ok = False

    for attempt in range(1, MAX_UPLOAD_RETRIES + 1):
        # Re-locate the file input on every attempt (SAPUI5 may have replaced the DOM node)
        file_input = frame.locator("#application-Nis-manage-component---CreateNIS--CreateHdrFileUploader-fu")
        if file_input.count() == 0:
            file_input = form.locator("input[type='file'][id$='-fu']").first
        file_input.wait_for(state="attached", timeout=timeout)

        file_input.set_input_files(str(p))

        # Give SAPUI5 a moment to process the change event and update the UI
        try:
            page.wait_for_timeout(1000)
        except Exception:
            pass

        # ── Check 1: re-query and inspect the current file input's .files ───
        fresh_input = frame.locator("#application-Nis-manage-component---CreateNIS--CreateHdrFileUploader-fu")
        if fresh_input.count() == 0:
            fresh_input = form.locator("input[type='file'][id$='-fu']").first

        attached_count = 0
        try:
            attached_count = fresh_input.evaluate("el => el.files ? el.files.length : 0")
        except Exception:
            pass

        if attached_count >= 1:
            upload_ok = True
            print(f"[UploadPDF] Attempt {attempt}: input.files confirmed ({attached_count} file(s)).")
            break

        # ── Check 2 (authoritative): HdrAttachmentInput text field has the filename ──
        attachment_input = frame.locator(
            "#application-Nis-manage-component---CreateNIS--HdrAttachmentInput-inner"
        )
        if attachment_input.count() > 0:
            try:
                expect(attachment_input).to_have_value(
                    re.compile(re.escape(p.name), re.I), timeout=5000
                )
                upload_ok = True
                print(f"[UploadPDF] Attempt {attempt}: HdrAttachmentInput confirmed filename '{p.name}'.")
                break
            except Exception:
                pass

        # ── Check 3: accessibility description ──────────────────────────────
        acc_descr = frame.locator("#application-Nis-manage-component---CreateNIS--CreateHdrFileUploader-AccDescr")
        if acc_descr.count() == 0:
            acc_descr = form.locator("[id$='-AccDescr']").first
        try:
            expect(acc_descr).to_contain_text(re.escape(p.name), timeout=3000)
            upload_ok = True
            print(f"[UploadPDF] Attempt {attempt}: AccDescr confirmed filename '{p.name}'.")
            break
        except Exception:
            pass

        print(f"[UploadPDF] Attempt {attempt}/{MAX_UPLOAD_RETRIES} — file not yet confirmed, retrying…")

    if not upload_ok:
        # Final best-effort: manually fire a change event on the current input
        try:
            fresh_input = frame.locator("#application-Nis-manage-component---CreateNIS--CreateHdrFileUploader-fu")
            if fresh_input.count() == 0:
                fresh_input = form.locator("input[type='file'][id$='-fu']").first
            fresh_input.evaluate("""
                (el) => {
                  const evt = new Event('change', { bubbles: true });
                  el.dispatchEvent(evt);
                }
            """)
            page.wait_for_timeout(1500)
        except Exception:
            pass
        # One last filename check before giving up
        attachment_input = frame.locator(
            "#application-Nis-manage-component---CreateNIS--HdrAttachmentInput-inner"
        )
        if attachment_input.count() > 0:
            try:
                expect(attachment_input).to_have_value(
                    re.compile(re.escape(p.name), re.I), timeout=5000
                )
                upload_ok = True
                print(f"[UploadPDF] Final fallback: HdrAttachmentInput confirmed '{p.name}'.")
            except Exception:
                pass

    if not upload_ok:
        print(f"[UploadPDF] WARNING — could not confirm file attachment after {MAX_UPLOAD_RETRIES} attempts. Proceeding anyway.")

    # This FileUploader auto-processes on selection (no separate Upload button in
    # the DOM), but click one if present for older variants.
    if click_upload_button:
        upload_btn = form.locator("button[id$='-uploadButton'], button[aria-label*='Upload']")
        if upload_btn.count() > 0:
            try:
                expect(upload_btn.first).to_be_visible(timeout=3000)
                upload_btn.first.click()
            except Exception:
                pass

    if wait_for_post:
        try:
            page.wait_for_timeout(1500)
        except Exception:
            pass

    print(f"✅ Header PDF selected: {p.name}")


def click_next_button3(frame, timeout: int = 15000):
    """
    Click the 'Next' button inside the NIS Create iframe.
    """
    next_btn = frame.get_by_role("button", name=re.compile(r"^\s*Next\s*$", re.I))

    if next_btn.count() == 0:
        next_btn = frame.locator("button:has-text('Next')")

    if next_btn.count() == 0:
        next_btn = frame.locator("#__button3")

    expect(next_btn.first).to_be_visible(timeout=timeout)
    try:
        expect(next_btn.first).to_be_enabled(timeout=timeout)
    except:
        pass

    try:
        next_btn.first.scroll_into_view_if_needed()
    except:
        pass

    next_btn.first.click()
    print("✅ 'Next' button clicked.")


def click_process_button(page, frame, timeout: int = 20000):
    """
    Click the 'Process' button inside the NIS app iframe.
    """
    try:
        frame.locator(".sapUiBlockLayer, .sapUiLocalBusyIndicator").wait_for(state="detached", timeout=5000)
    except Exception:
        pass

    btn = frame.get_by_role("button", name=re.compile(r"^\s*Process\s*$", re.I))
    if btn.count() == 0:
        btn = frame.locator("button:has-text('Process'), [role='button']:has-text('Process')")
    if btn.count() == 0:
        btn = frame.locator("#__button2")

    expect(btn.first).to_be_visible(timeout=timeout)
    try:
        expect(btn.first).to_be_enabled(timeout=timeout)
    except Exception:
        pass

    try:
        btn.first.scroll_into_view_if_needed(timeout=2000)
    except Exception:
        pass

    try:
        frame.locator(".sapUiBlockLayer").wait_for(state="detached", timeout=1500)
    except Exception:
        pass

    btn.first.click()

    try:
        page.wait_for_timeout(300)
    except Exception:
        pass

    print("✅ 'Process' button clicked.")


def fill_invoice_date(frame, date_value: str = "10.02.2020", timeout: int = 20000):
    """
    Fill the UI5 DatePicker 'Invoice Date' with the provided value.
    """
    date_input = frame.locator("#application-Nis-manage-component---CreateNIS--invoiceDate-inner")
    expect(date_input).to_be_visible(timeout=timeout)

    date_input.click()
    date_input.fill("")
    date_input.type(date_value, delay=20)
    date_input.press("Enter")

    expect(date_input).to_have_value(date_value, timeout=timeout)
    print(f"✅ Invoice Date filled: {date_value}")


def fill_vendor_bank_invoice(frame, vendor="215944", bank="59", invoice="AGE25CL05110128", timeout=20000):
    vendor_input = frame.locator("#application-Nis-manage-component---CreateNIS--vendorInput-inner")
    expect(vendor_input).to_be_visible(timeout=timeout)
    vendor_input.fill("")
    vendor_input.type(vendor, delay=30)
    vendor_input.press("Enter")
    print(f"✅ Vendor entered: {vendor}")

    bank_input = frame.locator("#application-Nis-manage-component---CreateNIS--bankInput-inner")
    expect(bank_input).to_be_visible(timeout=timeout)
    bank_input.fill("")
    bank_input.type(bank, delay=30)
    bank_input.press("Enter")
    print(f"✅ Bank entered: {bank}")

    inv_input = frame.locator("#application-Nis-manage-component---CreateNIS--invoiceNumberInput-inner")
    expect(inv_input).to_be_visible(timeout=timeout)
    inv_input.fill("")
    inv_input.type(invoice, delay=20)
    print(f"✅ Invoice entered: {invoice}")


def ensure_switch_on(page, frame, timeout: int = 15000):
    """
    Ensure the 'Dig. Sign Document' UI5 Switch is turned ON.

    Uses stable role/label selectors instead of the dynamic SAPUI5
    auto-generated IDs (#__switch0-*) which change on every page render.
    """
    # Primary: find by role + accessible label (most stable)
    sw = frame.get_by_role("switch", name=re.compile(r"Dig\.?\s*Sign", re.I))

    # Fallback 1: any switch whose label text contains "Dig"
    if sw.count() == 0:
        sw = frame.locator("[role='switch']").filter(
            has=frame.locator(":text('Dig')")
        )

    # Fallback 2: any switch inside the Invoice details form area
    if sw.count() == 0:
        sw = frame.locator("[role='switch']").first

    expect(sw.first).to_be_visible(timeout=timeout)

    # Determine current state via aria-checked attribute
    checked = ""
    try:
        checked = (sw.first.get_attribute("aria-checked") or "").strip().lower()
    except Exception:
        pass

    # Also check data-sap-ui-swt on any child element for SAP-specific state
    if checked not in ("true", "false"):
        try:
            handle = sw.first.locator("[data-sap-ui-swt]")
            if handle.count() > 0:
                swt = (handle.first.get_attribute("data-sap-ui-swt") or "").strip().lower()
                checked = "true" if swt == "on" else "false"
        except Exception:
            pass

    if checked != "true":
        sw.first.click()
        try:
            page.wait_for_timeout(300)
        except Exception:
            pass

    # Confirm it is ON
    try:
        final_checked = (sw.first.get_attribute("aria-checked") or "").strip().lower()
        assert final_checked == "true", f"Switch did not turn ON, aria-checked='{final_checked}'"
    except AssertionError:
        raise
    except Exception:
        pass  # If we can't re-read the attribute, trust the click happened

    print("✅ Switch is ON")



CATEGORY6_LABEL = "CATEGORY 6 - FINANCE/LEGAL & SEC/OTHER FUNC / ACTIVITY RELATED EXPS"

_CAT_COMBO_INPUT  = "#application-Nis-manage-component---CreateNIS--expCategoryCombo-inner"
_CAT_COMBO_ARROW  = "#application-Nis-manage-component---CreateNIS--expCategoryCombo-arrow"
_CAT_COMBO_POPUP  = "#application-Nis-manage-component---CreateNIS--expCategoryCombo-popup"
_SUB_COMBO_INPUT  = "#application-Nis-manage-component---CreateNIS--expSubCategoryCombo-inner"
_SUB_COMBO_ARROW  = "#application-Nis-manage-component---CreateNIS--expSubCategoryCombo-arrow"
_SUB_COMBO_POPUP  = "#application-Nis-manage-component---CreateNIS--expSubCategoryCombo-popup"

# ── Expense Description (Sub-Category) mapping per charge type ───────────────
# Each charge type must select its respective dropdown item:
#   DSM Charges      → "QCA PAYMENT ON ACCOUNT OF DSM - DEVIATION SETTLEMENT MECHANISM CHARGES" (GL 53065790)
#   Drawl Charges    → "QCA PAYMENT ON ACCOUNT OF DSM - CONSUMPTION OF DRAWL POWER"             (GL 53065780)
#   Revenue Diff     → "QCA PAYMENT ON ACCOUNT OF DSM"                                          (GL 53022100)
#   Revenue Loss     → "QCA PAYMENT ON ACCOUNT OF DSM"                                          (GL 53022100)
EXPENSE_SUBCATEGORY_MAP = {
    "total_dsm_charges_payable": "QCA PAYMENT ON ACCOUNT OF DSM - DEVIATION SETTLEMENT MECHANISM CHARGES",
    "drawl_charges_payable":     "QCA PAYMENT ON ACCOUNT OF DSM - CONSUMPTION OF DRAWL POWER",
    "revenue_diff":              "QCA PAYMENT ON ACCOUNT OF DSM",
    "revenue_loss":              "QCA PAYMENT ON ACCOUNT OF DSM",
}

# Default fallback when the charge type is not in the mapping
_DEFAULT_EXPENSE_SUBCATEGORY = "QCA PAYMENT ON ACCOUNT OF DSM"


def _close_any_open_popup(page, frame, timeout=3000):
    """
    If any SAP ComboBox popup / dialog is still open, close it with Escape.
    """
    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(400)
    except Exception:
        pass
    for sel in [".sapMComboBoxBasePicker", ".sapMDialog", ".sapMPopover"]:
        try:
            frame.locator(sel).wait_for(state="detached", timeout=timeout)
        except Exception:
            pass


def _dismiss_sap_ui_static_overlay(page, frame):
    """
    KEY FIX: After closing the Add Expense Item dialog (OK button), SAP UI5 sometimes
    leaves a ghost element inside #sap-ui-static that continues to intercept pointer
    events on the main page.  This helper waits for that overlay to detach / become
    hidden before proceeding.

    The error log shows:
      <input id="application-Nis-manage-component---CreateNIS--WBS-inner"/>
      from <div id="sap-ui-static"> subtree intercepts pointer events

    This means the "Add Item" dialog's DOM is still live inside #sap-ui-static even
    though it visually closed.  We force it out by:
      1) Pressing Escape to close any lingering dialog.
      2) Waiting for the WBS input (which only exists inside the Add Item dialog) to
         detach from the DOM.
      3) Waiting for the sap-ui-static subtree to stop containing visible dialogs.
      4) Pausing briefly for SAP UI5 to finish its cleanup animations.
    """
    import time

    # 1) Send Escape in case the dialog is still technically open
    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(300)
    except Exception:
        pass

    # 2) Wait for the WBS input (lives only inside the Add Item dialog) to detach
    wbs_selector = "#application-Nis-manage-component---CreateNIS--WBS-inner"
    try:
        frame.locator(wbs_selector).wait_for(state="detached", timeout=8000)
        print("[StaticOverlay] WBS input detached — dialog fully closed.")
    except Exception:
        # If WBS input doesn't detach in time, try a JS hide on #sap-ui-static children
        try:
            page.evaluate("""
                () => {
                    var staticArea = document.getElementById('sap-ui-static');
                    if (staticArea) {
                        var dialogs = staticArea.querySelectorAll(
                            '[role="dialog"], [role="alertdialog"], .sapMDialog, .sapMPopover'
                        );
                        dialogs.forEach(function(d) {
                            d.style.visibility = 'hidden';
                            d.style.pointerEvents = 'none';
                        });
                    }
                }
            """)
            print("[StaticOverlay] Forced #sap-ui-static dialogs to pointer-events:none via JS.")
        except Exception as js_err:
            print(f"[StaticOverlay] JS fallback also failed: {js_err}")

    # 3) Also wait for any general block-layer / busy indicator to clear
    for sel in [".sapUiBlockLayer", ".sapUiLocalBusyIndicator", ".sapMBusyDialog"]:
        try:
            frame.locator(sel).wait_for(state="detached", timeout=3000)
        except Exception:
            pass

    # 4) Final short settle to allow SAP UI5 re-render
    time.sleep(1)


def type_expense_category_label(page, frame, label_text: str = CATEGORY6_LABEL, timeout: int = 25000):
    """
    Select 'Expense Category' from the UI5 ComboBox.

    Always selects by TEXT match — never by position (aria-posinset) or dynamic
    IDs (#__item*) which can map to the wrong category when the list order changes.

    Strategy:
      1) Type "CATEGORY 6" to filter the dropdown, then click the matching item.
      2) JS SAP UI5 API — iterate items and select by text.
      3) Fill the full label text directly + Enter as last resort.
    """
    combo_input = frame.locator(_CAT_COMBO_INPUT)
    expect(combo_input).to_be_visible(timeout=timeout)

    selected = False

    # ── Approach 1: type-filter the dropdown so only CATEGORY 6 items show ──
    try:
        combo_input.click()
        page.wait_for_timeout(300)
        combo_input.fill("")
        combo_input.type("CATEGORY 6", delay=30)
        page.wait_for_timeout(1200)

        popup = frame.locator(f"{_CAT_COMBO_POPUP}, .sapMComboBoxBasePicker")
        try:
            expect(popup.first).to_be_visible(timeout=5000)
        except Exception:
            pass

        if popup.count() > 0:
            # Only select items whose text contains "CATEGORY 6" — never by position
            item = popup.locator(
                "[role='option']:has-text('CATEGORY 6'), "
                "li:has-text('CATEGORY 6'), "
                ".sapMLIB:has-text('CATEGORY 6'), "
                ".sapMComboBoxItem:has-text('CATEGORY 6')"
            )
            if item.count() == 0:
                item = popup.get_by_text(re.compile(r"CATEGORY\s*6", re.I))

            if item.count() > 0:
                item.first.click()
                page.wait_for_timeout(600)
                selected = True
                print(f"✅ Expense Category selected via type-filter: {label_text}")
            else:
                # Close dropdown without selecting — fall through to JS approach
                page.keyboard.press("Escape")
                page.wait_for_timeout(400)
    except Exception as e:
        print(f"[ExpCategory] Type-filter approach failed: {e}")
        try:
            page.keyboard.press("Escape")
        except Exception:
            pass

    # ── Approach 2: SAP UI5 JS API — set by item text ───────────────────────
    if not selected:
        try:
            combo_input.evaluate("""
                (el) => {
                    var sId = el.id.replace('-inner', '');
                    if (typeof sap !== 'undefined' && sap.ui && sap.ui.getCore) {
                        var oControl = sap.ui.getCore().byId(sId);
                        if (oControl) {
                            var items = oControl.getItems ? oControl.getItems() : [];
                            for (var i = 0; i < items.length; i++) {
                                var txt = items[i].getText ? items[i].getText() : '';
                                if (/CATEGORY\\s*6/i.test(txt)) {
                                    oControl.setSelectedItem(items[i]);
                                    if (oControl.fireChange) oControl.fireChange({ value: txt });
                                    if (oControl.fireSelectionChange) {
                                        oControl.fireSelectionChange({ selectedItem: items[i] });
                                    }
                                    return txt;
                                }
                            }
                        }
                    }
                    return '';
                }
            """)
            page.wait_for_timeout(600)
            # Verify JS actually set a CATEGORY 6 value
            current_val = combo_input.input_value() or ""
            if re.search(r"CATEGORY\s*6", current_val, re.I):
                selected = True
                print(f"✅ Expense Category set via SAP UI5 JS API: {current_val}")
        except Exception as e:
            print(f"[ExpCategory] JS API approach failed: {e}")

    # ── Approach 3: direct fill + Enter ─────────────────────────────────────
    if not selected:
        try:
            combo_input.click()
            combo_input.fill("")
            combo_input.type(label_text, delay=15)
            combo_input.press("Enter")
            page.wait_for_timeout(400)
            selected = True
            print(f"✅ Expense Category typed and committed directly: {label_text}")
        except Exception as e:
            print(f"[ExpCategory] Direct fill approach failed: {e}")

    try:
        frame.locator(_CAT_COMBO_POPUP).wait_for(state="detached", timeout=4000)
    except Exception:
        pass
    page.wait_for_timeout(800)

    # Validate — must contain CATEGORY 6
    try:
        expect(combo_input).to_have_value(
            re.compile(re.escape(label_text), re.I), timeout=timeout
        )
    except Exception:
        # Fallback: ensure the field contains "CATEGORY 6", not some other category
        _current = combo_input.input_value() or ""
        if not re.search(r"CATEGORY\s*6", _current, re.I):
            expect(combo_input).to_have_value(
                re.compile(r"CATEGORY\s*6", re.I), timeout=5000
            )

    print(f"✅ Expense Category confirmed: {label_text}")


def type_expense_subcategory(page, frame, value="QCA PAYMENT ON ACCOUNT OF DSM", timeout: int = 20000):
    """
    Select the Sub-Category from the UI5 ComboBox.

    IMPORTANT: The dropdown contains multiple items starting with "QCA PAYMENT ON ACCOUNT OF DSM".
    We must match the EXACT value text, not just any item containing "QCA PAYMENT".
    """
    sub_input = frame.locator(_SUB_COMBO_INPUT)
    expect(sub_input).to_be_visible(timeout=timeout)

    page.wait_for_timeout(500)

    selected = False

    # ── Helper: build a unique search prefix to narrow results ────────────────
    # All three variants start with "QCA PAYMENT ON ACCOUNT OF DSM" so we use
    # more specific text to narrow the dropdown filter when possible.
    if "DEVIATION" in value.upper():
        _search_text = "QCA PAYMENT ON ACCOUNT OF DSM - DEV"
    elif "DRAWL" in value.upper() or "CONSUMPTION" in value.upper():
        _search_text = "QCA PAYMENT ON ACCOUNT OF DSM - CON"
    else:
        _search_text = "QCA PAYMENT ON ACCOUNT OF DSM"

    arrow_btn = frame.locator(_SUB_COMBO_ARROW)
    if arrow_btn.count() > 0:
        try:
            sub_input.click()
            page.wait_for_timeout(300)
            arrow_btn.click()
            page.wait_for_timeout(1500)

            popup = frame.locator(
                f"{_SUB_COMBO_POPUP}, "
                ".sapMComboBoxBasePicker"
            )
            try:
                expect(popup.first).to_be_visible(timeout=5000)
            except Exception:
                pass

            if popup.count() > 0:
                search_input = popup.locator(
                    "input[type='search'], input[type='text'], "
                    ".sapMSearchFieldInner input, .sapMSFI input"
                ).first
                if search_input.count() > 0:
                    try:
                        search_input.click()
                        search_input.fill("")
                        search_input.type(_search_text, delay=30)
                        page.wait_for_timeout(800)
                    except Exception:
                        pass

                # Match by EXACT text — use the full value for precise matching
                item = popup.locator(f"[role='option']").filter(has_text=re.compile(
                    r"^\s*" + re.escape(value) + r"\s*$", re.I
                ))
                if item.count() == 0:
                    # Fallback: match items that contain the full value text
                    item = popup.locator(f"li:has-text('{value}')")
                if item.count() == 0:
                    item = popup.get_by_text(value, exact=False)

                if item.count() > 0:
                    item.first.click()
                    page.wait_for_timeout(500)
                    selected = True
                    print(f"✅ Expense Sub-Category selected from dialog picker: {value}")
                else:
                    page.keyboard.press("Escape")
                    page.wait_for_timeout(400)
        except Exception as e:
            print(f"[ExpSubCategory] Dialog picker approach failed: {e}")

    if not selected:
        try:
            sub_input.click()
            page.wait_for_timeout(400)

            focused_id = page.evaluate(
                "() => { try { return document.activeElement ? "
                "(document.activeElement.shadowRoot "
                "? document.activeElement.shadowRoot.activeElement : document.activeElement).id "
                ": '' } catch(e) { return '' } }"
            )
            if "expCategoryCombo" in str(focused_id):
                print("[ExpSubCategory] Focus was stolen by category combo — clicking sub-category input again")
                page.keyboard.press("Escape")
                page.wait_for_timeout(400)
                sub_input.click()
                page.wait_for_timeout(400)

            sub_input.fill("")
            sub_input.type(_search_text, delay=50)
            page.wait_for_timeout(1000)

            popup = frame.locator(f"{_SUB_COMBO_POPUP}, .sapMComboBoxBasePicker")
            # Match by exact value text — critical to distinguish between the 3 variants
            listbox_item = popup.locator("[role='option']").filter(has_text=re.compile(
                r"^\s*" + re.escape(value) + r"\s*$", re.I
            ))
            if listbox_item.count() == 0:
                listbox_item = popup.locator(f"li:has-text('{value}')")
            if listbox_item.count() == 0:
                listbox_item = frame.locator(f"[role='option']:has-text('{value}')")

            if listbox_item.count() > 0:
                listbox_item.first.click()
                page.wait_for_timeout(500)
                selected = True
                print(f"✅ Expense Sub-Category selected from type-ahead: {value}")
            else:
                sub_input.fill("")
                sub_input.type(value, delay=15)
                sub_input.press("Enter")
                page.wait_for_timeout(300)
                selected = True
                print(f"✅ Expense Sub-Category typed and committed: {value}")
        except Exception as e:
            print(f"[ExpSubCategory] Type-ahead approach failed: {e}")

    if not selected:
        print("[ExpSubCategory] All UI approaches failed, forcing value via JS...")
        sub_input.evaluate("""
            (el, val) => {
                var nativeInputValueSetter = Object.getOwnPropertyDescriptor(
                    window.HTMLInputElement.prototype, 'value'
                ).set;
                nativeInputValueSetter.call(el, val);
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
                var sId = el.id.replace('-inner', '');
                if (typeof sap !== 'undefined' && sap.ui && sap.ui.getCore) {
                    var oControl = sap.ui.getCore().byId(sId);
                    if (oControl) {
                        var items = oControl.getItems ? oControl.getItems() : [];
                        for (var i = 0; i < items.length; i++) {
                            var txt = items[i].getText ? items[i].getText() : '';
                            if (txt.toUpperCase() === val.toUpperCase()) {
                                oControl.setSelectedItem(items[i]);
                                if (oControl.fireChange) oControl.fireChange({ value: txt });
                                if (oControl.fireSelectionChange) {
                                    oControl.fireSelectionChange({ selectedItem: items[i] });
                                }
                                break;
                            }
                        }
                        if (oControl.getValue && oControl.getValue().toUpperCase() !== val.toUpperCase()) {
                            oControl.setValue(val);
                            if (oControl.fireChange) oControl.fireChange({ value: val });
                        }
                    }
                }
            }
        """, value)
        sub_input.press("Enter")
        print(f"✅ Expense Sub-Category forced via JS: {value}")

    try:
        frame.locator(_SUB_COMBO_POPUP).wait_for(state="detached", timeout=3000)
    except Exception:
        pass

    try:
        expect(sub_input).to_have_value(
            re.compile(re.escape(value), re.I), timeout=timeout
        )
    except Exception:
        expect(sub_input).to_have_value(
            re.compile(r"QCA", re.I), timeout=timeout
        )

    print(f"✅ Expense Sub-Category confirmed: {value}")


def fill_amount_value(frame, amount="4923", timeout: int = 15000):
    """
    Fill the 'Total Gross Value of Invoice (Including GST)' amount field.
    """
    amt_input = frame.locator("#application-Nis-manage-component---CreateNIS--Amount-inner")
    expect(amt_input).to_be_visible(timeout=timeout)

    amt_input.click()
    amt_input.fill("")
    amt_input.type(str(amount), delay=20)

    try:
        amt_input.press("Enter")
    except:
        pass

    expect(amt_input).to_have_value(str(amount), timeout=timeout)
    print(f"✅ Amount field filled: {amount}")


def fill_cost_center(frame, value="6067OPC1", timeout=20000):
    cc_input = frame.locator("#application-Nis-manage-component---CreateNIS--CostCenter-inner")
    expect(cc_input).to_be_visible(timeout=timeout)

    cc_input.click()
    cc_input.fill("")
    cc_input.type(value, delay=30)
    cc_input.press("Enter")

    expect(cc_input).to_have_value(value)
    print(f"✅ Cost Center filled: {value}")


def fill_plant(frame, value="6067", timeout=20000):
    plant_input = frame.locator("#application-Nis-manage-component---CreateNIS--Plant-inner")
    expect(plant_input).to_be_visible(timeout=timeout)

    plant_input.click()
    plant_input.fill("")
    plant_input.type(value, delay=30)
    plant_input.press("Enter")

    expect(plant_input).to_have_value(value)
    print(f"✅ Plant filled: {value}")


def _unique_visible_textarea(frame, css_selector: str, timeout=10000):
    """Return a single visible textarea handle for the given selector."""
    loc = frame.locator(css_selector).first
    expect(loc).to_be_visible(timeout=timeout)
    return loc


def fill_remarks(frame, value="2025.01.05 to 2025.01.11", timeout=20000):
    remarks = _unique_visible_textarea(
        frame, "#application-Nis-manage-component---CreateNIS--Remarks-inner", timeout=timeout
    )
    remarks.scroll_into_view_if_needed()
    remarks.click()

    remarks.evaluate("""
        (el, val) => {
            el.value = "";
            el.dispatchEvent(new Event("input", { bubbles: true }));
            el.value = String(val);
            el.dispatchEvent(new Event("input", { bubbles: true }));
            el.dispatchEvent(new Event("change", { bubbles: true }));
            el.blur();
        }
    """, value)

    try:
        remarks.press("Tab")
    except Exception:
        pass

    expect(remarks).to_have_js_property("value", value, timeout=timeout)

    actual = remarks.evaluate("""
        el => (el.value ?? "")
              .replace(/\\u00A0/g, " ")
              .trim()
              .replace(/\\s+/g, " ")
    """)
    expected = value.strip().replace("\u00A0", " ")
    if actual != expected:
        raise AssertionError(f"Remarks mismatch.\nExpected: '{expected}'\nActual:   '{actual}'")

    print(f"✅ Remarks filled: {value}")


def fill_additional_remarks(frame, value="pss8", timeout=20000):
    add_rem = _unique_visible_textarea(frame, "#__area0-inner", timeout=timeout)
    add_rem.scroll_into_view_if_needed()
    add_rem.click()

    add_rem.evaluate("""
        (el, val) => {
            el.value = "";
            el.dispatchEvent(new Event("input", { bubbles: true }));
            el.value = String(val);
            el.dispatchEvent(new Event("input", { bubbles: true }));
            el.dispatchEvent(new Event("change", { bubbles: true }));
            el.blur();
        }
    """, value)

    try:
        add_rem.press("Tab")
    except Exception:
        pass

    expect(add_rem).to_have_js_property("value", value, timeout=timeout)
    print(f"✅ Additional Remarks filled: {value}")


def click_acknowledge_checkbox(page, frame, timeout: int = 15000):
    """
    Tick the 'Acknowledge' checkbox inside the Add Expense Item dialog.

    SAPUI5 element:
      <input id="application-Nis-manage-component---CreateNIS--isAckKnowledgeChk-CB"
             type="CheckBox">
    Wrapper:
      <div  id="application-Nis-manage-component---CreateNIS--isAckKnowledgeChk-CbBg" …>
    """
    # Locate by the wrapper div (the visible click-target) — the raw <input> is
    # hidden behind it in SAP's CSS.
    cb_bg = frame.locator(
        "#application-Nis-manage-component---CreateNIS--isAckKnowledgeChk-CbBg"
    )
    if cb_bg.count() == 0:
        # Fallback: any checkbox whose id ends with isAckKnowledgeChk-CbBg
        cb_bg = frame.locator("div[id$='isAckKnowledgeChk-CbBg']")

    cb_input = frame.locator(
        "#application-Nis-manage-component---CreateNIS--isAckKnowledgeChk-CB"
    )
    if cb_input.count() == 0:
        cb_input = frame.locator("input[id$='isAckKnowledgeChk-CB']")

    # Wait for the checkbox to be attached (it lives inside the Add-Item dialog)
    try:
        cb_input.wait_for(state="attached", timeout=timeout)
    except Exception:
        print("[AckCheckbox] Checkbox not found in DOM — skipping.")
        return

    # Only click if not already checked
    is_checked = cb_input.is_checked()
    if not is_checked:
        try:
            cb_bg.first.scroll_into_view_if_needed()
        except Exception:
            pass
        cb_bg.first.click()
        try:
            page.wait_for_timeout(500)
        except Exception:
            pass

    # Verify
    try:
        expect(cb_input).to_be_checked(timeout=5000)
        print("✅ Acknowledge checkbox ticked.")
    except Exception:
        # JS fallback — force-check via the SAPUI5 control model
        try:
            cb_input.evaluate("""
                (el) => {
                    el.checked = true;
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                    var sId = el.id.replace('-CB', '');
                    if (typeof sap !== 'undefined' && sap.ui && sap.ui.getCore) {
                        var oCtrl = sap.ui.getCore().byId(sId);
                        if (oCtrl && oCtrl.setSelected) {
                            oCtrl.setSelected(true);
                            if (oCtrl.fireSelect) oCtrl.fireSelect({ selected: true });
                        }
                    }
                }
            """)
            print("✅ Acknowledge checkbox forced via JS.")
        except Exception as js_err:
            print(f"[AckCheckbox] WARNING — could not tick checkbox: {js_err}")


def click_ok_button(page, frame=None, timeout: int = 20000):
    """
    Click the 'OK' button on a SAPUI5 MessageBox/Dialog.
    """

    def _wait_blocklayer_gone(scope):
        try:
            scope.locator(".sapUiBlockLayer, .sapUiLocalBusyIndicator").wait_for(
                state="detached", timeout=3000
            )
        except Exception:
            pass

    def _try_click_in_scope(scope, scope_name="scope"):
        _wait_blocklayer_gone(scope)

        ok_btn = scope.locator("#__button18")
        print(f"[OK][{scope_name}] #__button18 count: {ok_btn.count()}")

        if ok_btn.count() == 0:
            ok_btn = scope.locator(
                ".sapMDialog button.sapMBtnAccept, "
                ".sapMMessageBox button.sapMBtnAccept, "
                ".sapMDialog .sapMBarChild bdi:text-is('OK'), "
                ".sapMMessageBox .sapMBarChild bdi:text-is('OK')"
            )
            print(f"[OK][{scope_name}] dialog footer OK count: {ok_btn.count()}")

        if ok_btn.count() == 0:
            ok_btn = scope.get_by_role("button", name=re.compile(r"^\s*OK\s*$", re.I))
            print(f"[OK][{scope_name}] role-based OK count: {ok_btn.count()}")

        if ok_btn.count() == 0:
            ok_btn = scope.locator("button:has(bdi:text-is('OK'))")
            print(f"[OK][{scope_name}] bdi text-is OK count: {ok_btn.count()}")

        if ok_btn.count() == 0:
            raise RuntimeError(f"[OK][{scope_name}] No OK button found.")

        expect(ok_btn.first).to_be_visible(timeout=timeout)

        try:
            ok_btn.first.scroll_into_view_if_needed()
        except Exception:
            pass

        _wait_blocklayer_gone(scope)

        try:
            ok_btn.first.click(timeout=5000)
            print(f"✅ OK clicked via normal click in {scope_name}")
            return True
        except Exception as click_err:
            print(f"[OK][{scope_name}] normal click failed: {click_err}, trying JS click...")

        ok_btn.first.evaluate("el => el.click()")
        print(f"✅ OK clicked via JS click in {scope_name}")
        return True

    if frame is not None:
        try:
            _try_click_in_scope(frame, scope_name="iframe")
            return
        except Exception as e_iframe:
            print(f"[OK][iframe] failed: {e_iframe}")

    try:
        _try_click_in_scope(page, scope_name="top-page")
        return
    except Exception as e_top:
        print(f"[OK][top-page] failed: {e_top}")

    print("[OK] All scoped attempts failed. Trying brute-force JS on page...")
    try:
        page.evaluate("""
            () => {
                let btn = document.querySelector('#__button18');
                if (btn) { btn.click(); return 'top'; }
                const iframes = document.querySelectorAll('iframe');
                for (const iframe of iframes) {
                    try {
                        btn = iframe.contentDocument.querySelector('#__button18');
                        if (btn) { btn.click(); return 'iframe'; }
                    } catch(e) { /* cross-origin */ }
                }
                throw new Error('__button18 not found anywhere');
            }
        """)
        print("✅ OK clicked via brute-force JS.")
        return
    except Exception as e_js:
        print(f"[OK] brute-force JS failed: {e_js}")

    raise RuntimeError("Could not locate/click the OK button (__button18) in page or any iframe.")


def click_ok(frame):
    """Click the OK button inside the dialog."""
    ok_button = frame.locator("#__mbox-btn-0")
    ok_button.click()
    print("OK button clicked")


def click_next(frame, timeout=15000):
    """
    Robust click for the SAPUI5 'Next' button after the Add Items screen.

    SAP UI5 auto-generates numeric button ids (e.g. __button6, __button10) that
    change every session / every render cycle, so we NEVER rely on a hardcoded id.

    Strategy (tried in order until one succeeds):
      1) ARIA role + exact accessible name "Next"  ← most stable
      2) <bdi> text content equals "Next"          ← works even when name attr is absent
      3) button whose visible text is exactly "Next"
      4) Any button whose id contains the data-ui5-accesskey="n" attribute
         (SAP uses accesskey="n" for the primary Next action)
      5) JS fallback: find every button in the iframe and click the one whose
         text is "Next" — bypasses any pointer-events overlay
    """
    next_btn = None

    # 1) ARIA role + name (preferred — survives id changes)
    candidate = frame.get_by_role("button", name=re.compile(r"^\s*Next\s*$", re.I))
    if candidate.count() > 0:
        next_btn = candidate

    # 2) <bdi> text match — SAP renders button labels inside <bdi> tags
    if next_btn is None or next_btn.count() == 0:
        candidate = frame.locator("button:has(bdi:text-is('Next'))")
        if candidate.count() > 0:
            next_btn = candidate

    # 3) Generic text-based button match
    if next_btn is None or next_btn.count() == 0:
        candidate = frame.locator("button:has-text('Next')")
        if candidate.count() > 0:
            next_btn = candidate

    # 4) SAP accesskey="n" — the primary Next button carries this attribute
    if next_btn is None or next_btn.count() == 0:
        candidate = frame.locator("button[data-ui5-accesskey='n']")
        if candidate.count() > 0:
            next_btn = candidate

    if next_btn is not None and next_btn.count() > 0:
        next_btn.first.wait_for(state="visible", timeout=timeout)
        try:
            next_btn.first.scroll_into_view_if_needed()
        except Exception:
            pass
        try:
            next_btn.first.click(timeout=timeout)
            print("✅ 'Next' button clicked successfully")
            return
        except Exception as e:
            print(f"[Next] Normal click failed ({e}), trying JS click...")
            try:
                next_btn.first.evaluate("el => el.click()")
                print("✅ 'Next' button clicked via JS click")
                return
            except Exception as js_e:
                print(f"[Next] JS element click also failed: {js_e}")

    # 5) Last-resort: brute-force JS — iterate all buttons inside the iframe
    #    and click the first one whose trimmed text is "Next"
    print("[Next] All locator strategies failed — trying brute-force JS button search...")
    try:
        frame.locator("body").evaluate("""
            () => {
                var buttons = document.querySelectorAll('button');
                for (var i = 0; i < buttons.length; i++) {
                    var txt = (buttons[i].innerText || buttons[i].textContent || '').trim();
                    if (txt === 'Next') {
                        buttons[i].click();
                        return 'clicked: ' + (buttons[i].id || 'no-id');
                    }
                }
                throw new Error('No Next button found in DOM');
            }
        """)
        print("✅ 'Next' button clicked via brute-force JS")
        return
    except Exception as e:
        raise RuntimeError(f"Could not find or click the 'Next' button: {e}")


def click_add_expense_item(page, frame, timeout=15000):
    """
    Robust click for the SAPUI5 'Add Expense Item' button.

    FIX: After clicking OK to close the previous Add Item dialog, SAP UI5 leaves a
    ghost WBS input inside #sap-ui-static that intercepts pointer events.  We must
    wait for that element to detach before attempting to click the Add button.
    This is done via _dismiss_sap_ui_static_overlay() which is called by the caller
    (add_one_item) but we also guard here as a safety net.
    """
    # Guard: wait for any lingering WBS input overlay to clear
    wbs_selector = "#application-Nis-manage-component---CreateNIS--WBS-inner"
    try:
        frame.locator(wbs_selector).wait_for(state="detached", timeout=5000)
    except Exception:
        # If still present, force pointer-events off via JS
        try:
            page.evaluate("""
                () => {
                    var el = document.getElementById('sap-ui-static');
                    if (el) {
                        el.querySelectorAll('[role="dialog"],[role="alertdialog"],.sapMDialog,.sapMPopover')
                          .forEach(function(d) {
                              d.style.pointerEvents = 'none';
                              d.style.visibility   = 'hidden';
                          });
                    }
                }
            """)
        except Exception:
            pass
        page.wait_for_timeout(500)

    # Also wait for any block layer
    try:
        frame.locator(".sapUiBlockLayer, .sapUiLocalBusyIndicator").wait_for(
            state="detached", timeout=3000
        )
    except Exception:
        pass

    # Try by ID first
    btn = frame.locator("#application-Nis-manage-component---CreateNIS--AddItemsButton")

    # Fallback: role + accessible name
    if btn.count() == 0:
        btn = frame.get_by_role("button", name="Add Expense Item")

    btn.wait_for(state="visible", timeout=timeout)

    # Try normal click first; fall back to JS click if intercepted
    try:
        btn.click(timeout=10000)
    except Exception as e:
        print(f"[AddItem] Normal click failed ({e}), trying JS click...")
        btn.evaluate("el => el.click()")

    print("✅ 'Add Expense Item' button clicked successfully")


def fill_payment_method(page, frame, value="NEFT/RTGS", timeout: int = 20000):
    """
    Fill the 'Payment Method' combobox with the provided value.
    """
    combo_input = frame.locator("#application-Nis-manage-component---CreateNIS--paymentMethodInput-inner")
    expect(combo_input).to_be_visible(timeout=timeout)

    selected = False

    arrow_btn = frame.locator("#application-Nis-manage-component---CreateNIS--paymentMethodInput-arrow")
    if arrow_btn.count() > 0:
        try:
            arrow_btn.click()
            page.wait_for_timeout(800)
            popup = frame.locator(
                "#application-Nis-manage-component---CreateNIS--paymentMethodInput-popup, "
                "[role='dialog'], .sapMDialog, .sapMPopover, .sapMComboBoxBasePicker, [role='listbox']"
            )
            try:
                expect(popup.first).to_be_visible(timeout=4000)
            except Exception:
                pass

            item = popup.locator(
                f"[role='option']:has-text('{value}'), "
                f".sapMComboBoxItem:has-text('{value}'), "
                f".sapMLIB:has-text('{value}')"
            )
            if item.count() > 0:
                item.first.click()
                page.wait_for_timeout(300)
                combo_input.press("Enter")
                selected = True
                print(f"✅ Payment Method selected from dialog: {value}")
        except Exception as e:
            print(f"[PaymentMethod] Dialog picker approach failed: {e}")

    if not selected:
        try:
            combo_input.click()
            combo_input.fill("")
            combo_input.type(value, delay=20)
            page.wait_for_timeout(600)

            list_item = frame.locator(
                f"[role='option']:has-text('{value}'), "
                f".sapMComboBoxItem:has-text('{value}'), "
                f".sapMLIB:has-text('{value}')"
            )
            if list_item.count() > 0:
                list_item.first.click()
                page.wait_for_timeout(200)
            else:
                combo_input.press("Enter")

            selected = True
            print(f"✅ Payment Method set via type-ahead: {value}")
        except Exception as e:
            print(f"[PaymentMethod] Type-ahead approach failed: {e}")

    if not selected:
        combo_input.evaluate(
            """
            (el, val) => {
                var setter = Object.getOwnPropertyDescriptor(
                    window.HTMLInputElement.prototype, 'value'
                ).set;
                setter.call(el, val);
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));

                var sId = el.id.replace('-inner', '');
                if (typeof sap !== 'undefined' && sap.ui && sap.ui.getCore) {
                    var oCtrl = sap.ui.getCore().byId(sId);
                    if (oCtrl) {
                        if (oCtrl.setValue) oCtrl.setValue(val);
                        if (oCtrl.fireChange) oCtrl.fireChange({ value: val });
                    }
                }
            }
            """,
            value,
        )
        combo_input.press("Enter")
        print(f"✅ Payment Method forced via JS: {value}")

    try:
        expect(combo_input).to_have_value(re.compile(re.escape(value), re.I), timeout=timeout)
    except Exception:
        actual = combo_input.evaluate("el => el.value || ''")
        assert value.lower() in actual.lower(), (
            f"Payment Method not set. Expected '{value}', got '{actual}'"
        )

    print(f"✅ Payment Method confirmed: {value}")


def click_approval_radio(page, frame, timeout: int = 20000):
    """
    Approval Final section: select the first radio button in the approver/fc table.

    IMPORTANT: After the Next click the approval screen may render inside a NEW
    iframe or even at the top-level page.  All scopes are searched:
      1) the existing frame locator passed in
      2) a freshly-resolved frame_locator (catches iframe reload after Next)
      3) every concrete Frame object in page.frames
      4) the top-level page itself
      5) page-level JS that walks ALL iframes (absolute last resort)

    The numeric prefix (__item17-) is auto-generated and changes every session,
    so we never use it.  We target by stable name attr, id suffix, SVG circle,
    wrapper class, then JS fallback.
    """
    import time

    iframe_selector = (
        "iframe[id^='application-Nis-manage-'], "
        "iframe[name^='application-Nis-manage-'], "
        "iframe[src*='Nis-manage']"
    )

    # Give the approval screen a moment to fully render after the Next click
    time.sleep(2)

    # ------------------------------------------------------------------
    # Helper: attempt every selector strategy in a given scope.
    # scope may be a FrameLocator, a Frame, or the Page itself.
    # Returns True on first success.
    # ------------------------------------------------------------------
    def _try_in_scope(scope, scope_label):
        candidates = [
            ("name attr",  "input[type='radio'][name*='fcTable_selectGroup']"),
            ("id suffix",  "input[type='radio'][id*='selectSingle-RB']"),
            ("SVG circle", "circle[id*='selectSingle-Button']"),
            ("sapMRbB",    ".sapMRbB"),
            ("any radio",  "input[type='radio']"),
        ]
        for label, sel in candidates:
            try:
                loc = scope.locator(sel)
                if loc.count() == 0:
                    continue
                # Attempt A: normal Playwright click
                try:
                    loc.first.click(timeout=4000)
                    print(f"[ApprovalRadio] clicked via [{label}] in {scope_label}")
                    return True
                except Exception:
                    pass
                # Attempt B: JS click (bypasses pointer-event blocks)
                try:
                    loc.first.evaluate(
                        "el => { el.click(); "
                        "el.dispatchEvent(new Event('change', {bubbles:true})); }"
                    )
                    print(f"[ApprovalRadio] JS-clicked via [{label}] in {scope_label}")
                    return True
                except Exception:
                    pass
            except Exception:
                continue
        return False

    selected = False

    # Scope 1: existing frame locator
    if not selected:
        selected = _try_in_scope(frame, "existing frame locator")

    # Scope 2: fresh frame_locator (catches iframe reload after Next)
    if not selected:
        try:
            fresh_frame = page.frame_locator(iframe_selector)
            selected = _try_in_scope(fresh_frame, "fresh frame_locator")
        except Exception as e:
            print(f"[ApprovalRadio] fresh frame_locator scope failed: {e}")

    # Scope 3: every concrete Frame object on the page
    if not selected:
        for i, f in enumerate(page.frames):
            try:
                if _try_in_scope(f, f"page.frames[{i}] ({f.url[:50]})"):
                    selected = True
                    break
            except Exception:
                continue

    # Scope 4: top-level page
    if not selected:
        selected = _try_in_scope(page, "top-level page")

    # Scope 5: page-level JS that walks every iframe document
    if not selected:
        print("[ApprovalRadio] All scoped strategies failed — trying page-level JS walk...")
        try:
            page.evaluate("""
                () => {
                    function tryClick(doc) {
                        var rb = doc.querySelector(
                            'input[type="radio"][name*="fcTable_selectGroup"]'
                        );
                        if (!rb) {
                            rb = doc.querySelector(
                                'input[type="radio"][id*="selectSingle"]'
                            );
                        }
                        if (!rb) {
                            rb = doc.querySelector('input[type="radio"]');
                        }
                        if (rb) {
                            rb.click();
                            rb.dispatchEvent(new Event('change', { bubbles: true }));
                            return true;
                        }
                        return false;
                    }
                    if (tryClick(document)) return 'top-level';
                    var iframes = document.querySelectorAll('iframe');
                    for (var i = 0; i < iframes.length; i++) {
                        try {
                            if (tryClick(iframes[i].contentDocument)) {
                                return 'iframe-' + i;
                            }
                        } catch (e) { /* cross-origin — skip */ }
                    }
                    throw new Error('No radio button found in any document');
                }
            """)
            selected = True
            print("✅ Approval radio selected via page-level JS iframe walk")
        except Exception as e:
            raise RuntimeError(
                "[ApprovalRadio] Could not select the approval radio button in any scope.\n"
                "Ensure the approval/fcTable screen has fully loaded before this step.\n"
                f"Last error: {e}"
            )

    try:
        page.wait_for_timeout(800)
    except Exception:
        pass
    print("✅ Approval radio step complete")


def _click_final_next(page, frame, timeout: int = 15000):
    """
    Click the final 'Next' button on the Approval screen.

    SAP UI5 auto-generates numeric button ids (e.g. __button10) that change
    every session, so we never rely on a hardcoded id.  Instead we try five
    strategies in order, including a page-level JS walk that crosses every
    iframe document — the same pattern used by click_approval_radio.

    The button HTML looks like:
        <button id="__button10" data-ui5-accesskey="n" ...>
            <bdi id="__button10-BDI-content">Next</bdi>
        </button>
    """
    iframe_selector = (
        "iframe[id^='application-Nis-manage-'], "
        "iframe[name^='application-Nis-manage-'], "
        "iframe[src*='Nis-manage']"
    )

    def _try_click(scope, label):
        """Try all locator strategies against *scope* (frame locator or Frame)."""
        strategies = [
            # 1) ARIA role + exact name (most stable)
            lambda s: s.get_by_role("button", name=re.compile(r"^\s*Next\s*$", re.I)),
            # 2) <bdi> inner text — SAP renders button labels inside <bdi>
            lambda s: s.locator("button:has(bdi:text-is('Next'))"),
            # 3) SAP accesskey="n" attribute carried by the primary Next button
            lambda s: s.locator("button[data-ui5-accesskey='n']"),
            # 4) Generic text fallback
            lambda s: s.locator("button:has-text('Next')"),
        ]
        for strategy in strategies:
            try:
                btn = strategy(scope)
                if btn.count() > 0:
                    btn.first.wait_for(state="visible", timeout=timeout)
                    try:
                        btn.first.scroll_into_view_if_needed()
                    except Exception:
                        pass
                    try:
                        btn.first.click(timeout=timeout)
                        print(f"✅ Final 'Next' clicked via locator [{label}]")
                        return True
                    except Exception:
                        # Pointer-event blocked — try JS click on the element
                        try:
                            btn.first.evaluate("el => el.click()")
                            print(f"✅ Final 'Next' clicked via JS element.click() [{label}]")
                            return True
                        except Exception:
                            pass
            except Exception:
                pass
        return False

    # Scope 1 — existing frame locator
    if _try_click(frame, "existing frame locator"):
        return

    # Scope 2 — fresh frame_locator (handles iframe reload after approval step)
    try:
        fresh_frame = page.frame_locator(iframe_selector)
        if _try_click(fresh_frame, "fresh frame_locator"):
            return
    except Exception:
        pass

    # Scope 3 — iterate every concrete Frame on the page
    for i, f in enumerate(page.frames):
        try:
            if _try_click(f, f"page.frames[{i}]"):
                return
        except Exception:
            continue

    # Scope 4 — top-level page
    if _try_click(page, "top-level page"):
        return

    # Scope 5 — JS walk across every iframe document (bypasses Playwright locators)
    print("[FinalNext] All locator strategies failed — trying JS iframe walk...")
    try:
        page.evaluate("""
            () => {
                function tryClick(doc) {
                    // Prefer buttons with data-ui5-accesskey="n"
                    var btn = doc.querySelector('button[data-ui5-accesskey="n"]');
                    if (!btn) {
                        // Fall back to any button whose visible text is "Next"
                        var all = doc.querySelectorAll('button');
                        for (var i = 0; i < all.length; i++) {
                            var txt = (all[i].innerText || all[i].textContent || '').trim();
                            if (txt === 'Next') { btn = all[i]; break; }
                        }
                    }
                    if (btn) {
                        btn.click();
                        return true;
                    }
                    return false;
                }
                if (tryClick(document)) return 'top-level';
                var iframes = document.querySelectorAll('iframe');
                for (var i = 0; i < iframes.length; i++) {
                    try {
                        if (tryClick(iframes[i].contentDocument)) return 'iframe-' + i;
                    } catch (e) { /* cross-origin — skip */ }
                }
                throw new Error('No Next button found in any document');
            }
        """)
        print("✅ Final 'Next' clicked via JS iframe walk")
        return
    except Exception as e:
        raise RuntimeError(
            "[FinalNext] Could not find or click the final 'Next' button in any scope.\n"
            "Ensure the Approval screen has fully loaded before this step.\n"
            f"Last error: {e}"
        )


def _click_fcTable_radio(page, frame, timeout: int = 10000):
    """
    Click the approver radio button in the fcTable on the Approval screen.

    SAP UI5 renders the radio as:
        <div class="sapMRbB sapMRbHoverable">          ← REAL click target
            <svg>...</svg>
            <input type="radio"
                   name="...-fcTable_selectGroup"      ← stable name
                   id="...-selectSingle-RB"            ← stable id suffix
                   tabindex="-1">                      ← NOT directly clickable
        </div>

    Key insight: the <input> has tabindex="-1" — SAP intentionally blocks direct
    clicks on it.  The correct target is the PARENT .sapMRbB wrapper div, which
    SAP's event handler listens to.  We locate the input by its stable name attr
    and then click its parent.

    Two attempts:
      1) Playwright: locate .sapMRbB that contains the named radio → click()
      2) JS fallback: querySelector by name, walk to parentElement, call .click()
         + fire 'change' event so SAP's model updates.
    """
    # Wait for the row to render after the Next click
    try:
        page.wait_for_timeout(2000)
    except Exception:
        pass

    # Locate the .sapMRbB wrapper that contains the fcTable radio input.
    # The name attribute is fully stable (no auto-generated prefix on it).
    wrapper = frame.locator(
        ".sapMRbB:has(input[type='radio'][name*='fcTable_selectGroup'])"
    )

    # Attempt 1: standard Playwright click on the wrapper div
    try:
        wrapper.first.wait_for(state="visible", timeout=timeout)
        wrapper.first.scroll_into_view_if_needed()
        wrapper.first.click(timeout=timeout)
        print("✅ [fcTableRadio] .sapMRbB wrapper clicked via Playwright")
        return
    except Exception as e:
        print(f"[fcTableRadio] Playwright wrapper click failed ({e}), trying JS fallback...")

    # Attempt 2: JS — find the input by name, click its parentElement (.sapMRbB),
    # then fire SAP's expected events on both parent and input
    try:
        frame.locator("body").evaluate("""
            () => {
                var input = document.querySelector(
                    'input[type="radio"][name*="fcTable_selectGroup"]'
                );
                if (!input) throw new Error('fcTable radio input not found');
                var wrapper = input.closest('.sapMRbB') || input.parentElement;
                if (!wrapper) throw new Error('sapMRbB wrapper not found');
                wrapper.click();
                input.checked = true;
                input.dispatchEvent(new Event('change', { bubbles: true }));
                wrapper.dispatchEvent(new Event('click',  { bubbles: true }));
            }
        """)
        print("✅ [fcTableRadio] .sapMRbB wrapper clicked via JS fallback")
        return
    except Exception as e:
        raise RuntimeError(
            "[fcTableRadio] Could not click the fcTable radio button.\n"
            "Ensure the Approval/approver screen is fully loaded.\n"
            f"Last error: {e}"
        )



def _dismiss_add_item_popup(page, frame=None, timeout: int = 4000):
    """
    After each 'Add Item' confirmation, the NIS portal may show an optional
    splash / info MessageBox with an OK button (id=__mbox-btn-0).
    This popup does NOT always appear — silently skip if absent.

    Button HTML the portal renders:
        <button id="__mbox-btn-0" ... class="sapMBtn ... sapMBtnEmphasized">
            <bdi id="__mbox-btn-0-BDI-content">OK</bdi>
        </button>

    Strategy (three attempts, all silent on miss):
      1. Look inside the iframe frame for #__mbox-btn-0
      2. Look in #sap-ui-static (overlay container that lives in the top-level DOM)
      3. Look on the top-level page
    """
    def _try_click(scope, scope_name: str) -> bool:
        try:
            btn = scope.locator("#__mbox-btn-0")
            if btn.count() == 0:
                # Also match by the BDI content in case the id is slightly different
                btn = scope.locator(
                    "button.sapMBtnEmphasized:has(bdi:text-is('OK')), "
                    "button[id^='__mbox-btn']:has(bdi:text-is('OK'))"
                )
            if btn.count() == 0:
                return False
            btn.first.wait_for(state="visible", timeout=timeout)
            try:
                btn.first.scroll_into_view_if_needed()
            except Exception:
                pass
            try:
                btn.first.click(timeout=3000)
            except Exception:
                btn.first.evaluate("el => el.click()")
            print(f"✅ [AddItemPopup] OK splash dismissed in {scope_name}")
            return True
        except Exception:
            return False

    # 1. Try inside the NIS iframe
    if frame is not None:
        if _try_click(frame, "iframe"):
            return

    # 2. Try the #sap-ui-static overlay (top-level DOM, outside iframe)
    try:
        static_layer = page.locator("#sap-ui-static")
        if static_layer.count() > 0:
            if _try_click(static_layer, "#sap-ui-static"):
                return
    except Exception:
        pass

    # 3. Try top-level page
    if _try_click(page, "top-page"):
        return

    # Popup did not appear — that is expected, nothing to do
    print("[AddItemPopup] No splash popup found after Add Item — continuing.")


def _dismiss_error_popup(page, frame=None, timeout: int = 5000):
    """
    Dismiss ALL SAP UI5 error/warning MessageBox popups that appear.

    These popups have ids like '__error0', '__error1' etc. with a footer
    containing a Close button:
        <button id="__mbox-btn-0" ...>
            <bdi id="__mbox-btn-0-BDI-content">Close</bdi>
        </button>

    Common error texts:
        - "Invalid resource path nisService.checkInvoiceSplit"
        - "Internal Server Error"
        - "Invalid resource path nisService.aiFeatureDisabled"

    Multiple popups can appear one after another — this function loops until
    all are dismissed. Safe to call anywhere — silently does nothing if no
    error popup is present.
    """
    import time as _time

    def _try_dismiss_one(scope, scope_name: str) -> bool:
        try:
            # Look for Close button in error/message dialogs
            close_btn = scope.locator(
                "button[id^='__mbox-btn']:has(bdi:text-is('Close')), "
                "button:has(bdi:text-is('Close')), "
                "[id^='__error'] button:has-text('Close'), "
                ".sapMMsgBox button:has-text('Close'), "
                ".sapMMessageDialog button:has-text('Close')"
            )
            if close_btn.count() == 0:
                return False

            close_btn.first.wait_for(state="visible", timeout=timeout)
            try:
                close_btn.first.click(timeout=3000)
            except Exception:
                close_btn.first.evaluate("el => el.click()")
            print(f"  ⚠️ [ErrorPopup] Error dialog dismissed in {scope_name}")
            return True
        except Exception:
            return False

    dismissed_count = 0
    max_popups = 10  # safety limit to avoid infinite loop

    for _ in range(max_popups):
        found = False

        # 1. Try the #sap-ui-static overlay (error dialogs usually render here)
        try:
            static_layer = page.locator("#sap-ui-static")
            if static_layer.count() > 0:
                if _try_dismiss_one(static_layer, "#sap-ui-static"):
                    found = True
                    dismissed_count += 1
                    _time.sleep(1)  # wait for next popup to appear
                    continue
        except Exception:
            pass

        # 2. Try top-level page
        if _try_dismiss_one(page, "top-page"):
            found = True
            dismissed_count += 1
            _time.sleep(1)
            continue

        # 3. Try inside the NIS iframe
        if frame is not None:
            if _try_dismiss_one(frame, "iframe"):
                found = True
                dismissed_count += 1
                _time.sleep(1)
                continue

        # No more popups found — done
        if not found:
            break

    if dismissed_count > 0:
        print(f"  ⚠️ [ErrorPopup] Total {dismissed_count} error popup(s) dismissed.")
    return dismissed_count > 0


def click_submit_button(page, frame, timeout: int = 20000):
    """
    Click the NIS 'Submit' button on the final confirmation screen.

    Button HTML:
        <button id="application-Nis-manage-component---CreateNIS--submitButton"
                data-ui5-accesskey="s" ...>
            <bdi id="application-Nis-manage-component---CreateNIS--submitButton-BDI-content">
                Submit
            </bdi>
        </button>

    Strategies tried in order (most specific → most generic):
      1. Exact element ID (most stable when the component id is fixed)
      2. SAP accesskey="s" attribute
      3. ARIA role + exact name "Submit"
      4. CSS :has(bdi) text match
      5. Generic has-text fallback
      6. JavaScript document walk across all iframe documents (last resort)
    """
    iframe_selector = (
        "iframe[id^='application-Nis-manage-'], "
        "iframe[name^='application-Nis-manage-'], "
        "iframe[src*='Nis-manage']"
    )

    SUBMIT_ID = "application-Nis-manage-component---CreateNIS--submitButton"

    def _attempt(scope, label: str) -> bool:
        strategies = [
            # 1) Exact element ID
            lambda s: s.locator(f"#{SUBMIT_ID}"),
            # 2) SAP accesskey attribute
            lambda s: s.locator("button[data-ui5-accesskey='s']"),
            # 3) ARIA role + name
            lambda s: s.get_by_role("button", name=re.compile(r"^\s*Submit\s*$", re.I)),
            # 4) <bdi> inner text
            lambda s: s.locator("button:has(bdi:text-is('Submit'))"),
            # 5) Generic has-text
            lambda s: s.locator("button:has-text('Submit')"),
        ]
        for strategy in strategies:
            try:
                btn = strategy(scope)
                if btn.count() == 0:
                    continue
                btn.first.wait_for(state="visible", timeout=timeout)
                try:
                    btn.first.scroll_into_view_if_needed()
                except Exception:
                    pass
                try:
                    btn.first.click(timeout=timeout)
                    print(f"✅ Submit button clicked [{label}]")
                    return True
                except Exception:
                    # Pointer-event blocked — JS click fallback
                    try:
                        btn.first.evaluate("el => el.click()")
                        print(f"✅ Submit button clicked via JS [{label}]")
                        return True
                    except Exception:
                        pass
            except Exception:
                pass
        return False

    # Scope 1 — existing frame locator
    if _attempt(frame, "existing frame locator"):
        return

    # Scope 2 — fresh frame_locator (handles iframe reload between steps)
    try:
        fresh_frame = page.frame_locator(iframe_selector)
        if _attempt(fresh_frame, "fresh frame_locator"):
            return
    except Exception:
        pass

    # Scope 3 — JS walk across every iframe document (cross-origin safe within same app)
    try:
        clicked = page.evaluate(f"""
            () => {{
                const SUBMIT_ID = "{SUBMIT_ID}";
                // Search in every frame document
                const frames = [window, ...Array.from(window.frames)];
                for (const f of frames) {{
                    try {{
                        let btn = f.document.getElementById(SUBMIT_ID);
                        if (!btn) {{
                            // Fallback: find by accesskey
                            btn = f.document.querySelector("button[data-ui5-accesskey='s']");
                        }}
                        if (!btn) {{
                            // Fallback: find by text content
                            const allBtns = f.document.querySelectorAll("button");
                            for (const b of allBtns) {{
                                if (b.textContent.trim().toLowerCase() === "submit") {{
                                    btn = b;
                                    break;
                                }}
                            }}
                        }}
                        if (btn) {{
                            btn.click();
                            return true;
                        }}
                    }} catch (e) {{}}
                }}
                return false;
            }}
        """)
        if clicked:
            print("✅ Submit button clicked via JS document walk.")
            return
    except Exception:
        pass

    print("⚠️  Submit button not found — page may have already submitted or selector changed.")


def book_nis(
    pdf_path: str,
    company_code: str,
    vendor_code: str,
    bank_key: str,
    invoice_date: str,
    cost_center: str,
    plant: str,
    purpose: str,
    svp_dsm_row: dict,
    max_internal_retries: int = 5,
    login_barrier=None,
    expense_items=None,
    invoice_number: str = "",
):
    """End-to-end NIS booking driven by data from main.py.

    Expected ``svp_dsm_row`` keys (positive values only will create items):
      - "total_dsm_charges_payable"
      - "drawl_charges_payable"
      - "revenue_diff"
      - "revenue_loss"

    ``expense_items`` (optional) overrides the per-field items derived from
    ``svp_dsm_row``. It is a list of dicts, one per expense line, used when a
    Wind/Solar SPV pair is booked together in a SINGLE NIS booking::

        {"amount": 123.0, "label": "...", "cost_center": "...",
         "plant": "...", "purpose": "...", "additional_remarks": "PSS8"}

    Each item keeps its own cost center / plant so both variants book correctly.

    The browser is launched ONCE and kept alive across internal retries.
    If a step fails mid-flow, the page navigates back to the Fiori launchpad
    and retries the entire booking form without closing/reopening the browser
    or re-authenticating.
    """
    with sync_playwright() as p:
        import time

        checklist_id = ""   # will be populated after final submission
        lm.log(f"NIS_BOOKING_START pdf={pdf_path} company={company_code} "
               f"vendor={vendor_code} bank={bank_key} invoice_date={invoice_date} "
               f"cost_center={cost_center} plant={plant}")

        try:
            browser = p.chromium.launch(
                executable_path=EDGE_PATH,
                headless=False,
                args=[
                    "--start-maximized",
                    "--disable-infobars",
                    "--force-device-scale-factor=1",
                ],
            )
        except Exception as exc:
            lm.log_error(f"NIS_BROWSER_LAUNCH_FAILED edge={EDGE_PATH}", exc=exc)
            raise

        # ── Authentication mode ───────────────────────────────────────────────
        auth_file   = AUTH_FILE
        use_cookies = getattr(config, "NIS_USE_COOKIES", False)

        try:
            if use_cookies:
                if not Path(auth_file).exists():
                    raise FileNotFoundError(
                        f"[Auth] NIS_USE_COOKIES is enabled but cookie file not found: {auth_file}\n"
                        "Run loginwithpassword.py separately and save the storage_state first."
                    )
                lm.log(f"NIS_AUTH_COOKIES file={auth_file}")
                print(f"[Auth] Loading existing storage_state from: {auth_file}")
                context = browser.new_context(viewport=None, storage_state=auth_file)
                page = context.new_page()
                print(f"[Auth] Navigating to NIS portal: {LOGIN_URL}")
                page.goto(LOGIN_URL, wait_until="networkidle")
            else:
                lm.log("NIS_AUTH_PASSWORD")
                context = browser.new_context(viewport=None)
                page = context.new_page()
                print("[Auth] Performing password-based login.")
                _login_with_password(page)
            lm.log("NIS_AUTH_OK")
        except Exception as exc:
            lm.log_error("NIS_AUTH_FAILED", exc=exc)
            # If this session was part of the synchronized first wave, abort the
            # barrier so the other sessions don't wait out the full timeout for a
            # login that will never arrive.
            if login_barrier is not None:
                try:
                    login_barrier.abort()
                except Exception:
                    pass
            browser.close()
            raise

        # Wait for the Fiori launchpad shell to be fully loaded
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(3000)
                  
        # ✅ Force website to fixed 500x500 (right & bottom free space)
        apply_viewport_fix(page)

        # ── Login barrier ─────────────────────────────────────────────────────
        # In parallel mode every session opens + logs in at the same time, then
        # waits here until ALL sessions have finished logging in. Only then does
        # any session start filling the booking form. A timeout (or a broken
        # barrier from a partial final batch) lets a session proceed on its own
        # so one stuck login can never hang the whole batch.
        if login_barrier is not None:
            try:
                print("[Login] ✅ Logged in — waiting for the other sessions to finish login …")
                login_barrier.wait()
                print("[Login] 🚦 All sessions logged in — starting to fill the form.")
            except Exception as bexc:
                print(f"[Login] Proceeding without full sync (barrier {type(bexc).__name__}).")

        # ── Internal retry loop — browser stays open across retries ───────────
        last_exc = None
        for attempt in range(1, max_internal_retries + 1):
            try:
                checklist_id = _do_booking_flow(
                    page, context, browser,
                    pdf_path=pdf_path,
                    company_code=company_code,
                    vendor_code=vendor_code,
                    bank_key=bank_key,
                    invoice_date=invoice_date,
                    cost_center=cost_center,
                    plant=plant,
                    purpose=purpose,
                    svp_dsm_row=svp_dsm_row,
                    attempt=attempt,
                    expense_items=expense_items,
                    invoice_number=invoice_number,
                )
                # Success — break out of retry loop
                break
            except Exception as exc:
                last_exc = exc
                lm.log_error(f"NIS_BOOKING_ATTEMPT_{attempt}_FAILED", exc=exc)
                print(f"\n  ⚠️  NIS booking attempt {attempt}/{max_internal_retries} failed: {exc}")

                if attempt >= max_internal_retries:
                    # All retries exhausted — close browser and propagate
                    break

                # ── Navigate back to launchpad home for next attempt ───────────
                print(f"  🔄 Navigating back to launchpad for retry (attempt {attempt + 1}) …")
                try:
                    page.goto(LOGIN_URL, wait_until="networkidle")
                    page.wait_for_timeout(3000)
                    apply_viewport_fix(page)
                    lm.log(f"NIS_RETRY_NAV_HOME attempt={attempt + 1}")
                except Exception as nav_exc:
                    # If navigation fails, try opening a fresh page in the same context
                    lm.log_error("NIS_RETRY_NAV_HOME_FAILED, opening fresh page", exc=nav_exc)
                    try:
                        page.close()
                    except Exception:
                        pass
                    page = context.new_page()
                    page.goto(LOGIN_URL, wait_until="networkidle")
                    page.wait_for_timeout(3000)
                    apply_viewport_fix(page)

                time.sleep(3)

        # Done — close browser
        page.wait_for_timeout(2000)
        browser.close()

        if checklist_id:
            lm.log(f"NIS_BOOKING_DONE checklist_id={checklist_id}")
            return checklist_id
        elif last_exc:
            lm.log_error(f"NIS_BOOKING_ALL_ATTEMPTS_FAILED after {max_internal_retries} attempts")
            raise last_exc
        else:
            lm.log(f"NIS_BOOKING_DONE checklist_id=n/a")
            return ""


def _do_booking_flow(
    page, context, browser,
    pdf_path: str,
    company_code: str,
    vendor_code: str,
    bank_key: str,
    invoice_date: str,
    cost_center: str,
    plant: str,
    purpose: str,
    svp_dsm_row: dict,
    attempt: int = 1,
    expense_items=None,
    invoice_number: str = "",
) -> str:
    """Execute the NIS booking form steps. Raises on failure (caller handles retry).

    Returns the NIS checklist ID on success, or "" if capture fails (non-fatal).
    """
    import time

    try:
        lm.log(f"NIS_OPEN_VIEW_CREATE attempt={attempt}")
        open_view_create(page)
        lm.log("NIS_OPEN_VIEW_CREATE_OK")
    except Exception as exc:
        lm.log_error("NIS_OPEN_VIEW_CREATE_FAILED", exc=exc)
        raise

    try:
        click_create_nis(page)
        print("7: clicked on the nis button")
    except Exception as exc:
        lm.log_error("NIS_CLICK_CREATE_NIS_FAILED", exc=exc)
        raise

    print("8: clicked on the nis button")
    frame = get_nis_frame(page)
    print("9: clicked on the nis button")

    try:
        close_failed_employee_dialog(page, timeout=8000)
        lm.log_warn("NIS_EMPLOYEE_DIALOG_CLOSED (appeared and was dismissed)")
    except Exception:
        pass  # dialog didn't appear — normal

    print("10: clicked on the nis button")
    # 1) Upload header PDF
    try:
        lm.log(f"NIS_UPLOAD_HEADER_PDF file={pdf_path}")
        upload_header_pdf(
            page,
            frame,
            pdf_path,
            click_upload_button=True,
            wait_for_post=True,
        )
        lm.log("NIS_UPLOAD_HEADER_PDF_OK")
    except Exception as exc:
        lm.log_error(f"NIS_UPLOAD_HEADER_PDF_FAILED file={pdf_path}", exc=exc)
        raise
    time.sleep(3)

    # Dismiss any error popup that may appear after PDF upload
    _dismiss_error_popup(page, frame)

    print("12: clicked on the nis button")

    # 2) Sub Document Type & Company Code
    try:
        lm.log("NIS_FILL_SUB_DOC_TYPE value=Non PO Vendor Payment")
        fill_sub_doc_type(page, frame, "Non PO Vendor Payment")
        lm.log("NIS_FILL_SUB_DOC_TYPE_OK")
    except Exception as exc:
        lm.log_error("NIS_FILL_SUB_DOC_TYPE_FAILED", exc=exc)
        raise
    time.sleep(2)

    print("company code : ", company_code)

    try:
        lm.log(f"NIS_FILL_COMPANY_CODE value={company_code}")
        fill_company_code(page, frame, company_code)
        lm.log("NIS_FILL_COMPANY_CODE_OK")
    except Exception as exc:
        lm.log_error(f"NIS_FILL_COMPANY_CODE_FAILED value={company_code}", exc=exc)
        raise
    time.sleep(2)

    # Dismiss any error popup before proceeding to Process
    _dismiss_error_popup(page, frame)

    # 3) Click Process
    try:
        lm.log("NIS_CLICK_PROCESS")
        click_process_button(page, frame)
        lm.log("NIS_CLICK_PROCESS_OK")
    except Exception as exc:
        lm.log_error("NIS_CLICK_PROCESS_FAILED", exc=exc)
        raise
    time.sleep(5)

    # Dismiss any error popup that may appear after Process click
    _dismiss_error_popup(page, frame)

    # -------- vendor & invoice section ---------

    def fill_vendor_field(frame_loc, vendor):
        """Fill vendor input; fallback to combined helper if needed."""
        try:
            v = frame_loc.locator("#application-Nis-manage-component---CreateNIS--vendorInput-inner")
            expect(v).to_be_visible(timeout=10000)
            v.fill("")
            v.type(str(vendor), delay=30)
            v.press("Enter")
        except Exception:
            try:
                fill_vendor_bank_invoice(frame_loc, vendor=str(vendor), bank=bank_key, invoice="")
            except Exception:
                pass

    def fill_bank_key_field(frame_loc, bank):
        """Fill bank key field only."""
        try:
            b = frame_loc.locator("#application-Nis-manage-component---CreateNIS--bankInput-inner")
            expect(b).to_be_visible(timeout=10000)
            b.fill("")
            b.type(str(bank), delay=30)
            b.press("Enter")
        except Exception:
            try:
                fill_vendor_bank_invoice(frame_loc, vendor=vendor_code, bank=str(bank), invoice="")
            except Exception:
                pass

    def fill_invoice_number_field(frame_loc, invoice_number):
        """Fill invoice number field."""
        try:
            inv = frame_loc.locator("#application-Nis-manage-component---CreateNIS--invoiceNumberInput-inner")
            expect(inv).to_be_visible(timeout=10000)
            inv.fill("")
            inv.type(str(invoice_number), delay=20)
        except Exception:
            try:
                fill_vendor_bank_invoice(frame_loc, vendor=vendor_code, bank=bank_key, invoice=str(invoice_number))
            except Exception:
                pass

    def fill_invoice_date_field(frame_loc, date_str):
        """Fill invoice date."""
        try:
            fill_invoice_date(frame_loc, date_value=str(date_str))
        except Exception:
            try:
                d = frame_loc.locator("#application-Nis-manage-component---CreateNIS--invoiceDate-inner")
                expect(d).to_be_visible(timeout=10000)
                d.fill("")
                d.type(str(date_str), delay=20)
                d.press("Enter")
            except Exception:
                pass

    import random
    # ── Invoice number selection ──────────────────────────────────────────────
    # For the special vendor codes the PDF carries no invoice data, so we keep
    # auto-generating the number (cost-center name + random digits, ≤15 chars).
    # For every other vendor we extract the invoice number from the attached PDF
    # (pdfplumber + LLM) and use that instead.
    cc_clean   = re.sub(r"[^A-Za-z0-9]", "", str(cost_center))
    rand_part  = "".join(str(random.randint(0, 9)) for _ in range(max(1, 15 - len(cc_clean))))
    generated_invoice_no = (cc_clean + rand_part)[:15]

    vendor_code_norm = str(vendor_code or "").strip()
    skip_extraction = getattr(config, "NIS_SKIP_INVOICE_EXTRACTION", True)
    provided_invoice = str(invoice_number or "").strip()

    # ── AWEK exception rule ──────────────────────────────────────────────────
    # SPVs matching NIS_EMAIL_PDF_SPV_KEYWORDS (e.g. AWEK06L, AWEK03L, AWEK05L)
    # must ALWAYS auto-generate the invoice number. They use email.pdf as the
    # upload document (handled in main.py), so the attached invoice PDF carries
    # no extractable invoice data. Force skip_extraction for them here.
    _spv_name_for_check = str((svp_dsm_row or {}).get("spv_name") or "").strip()
    _is_email_pdf_spv = is_email_pdf_spv(_spv_name_for_check)
    if _is_email_pdf_spv:
        skip_extraction = True
        provided_invoice = ""  # never use caller-provided extract for these SPVs
        lm.log(f"NIS_AWEK_EXCEPTION spv={_spv_name_for_check} → forcing auto-generated invoice")

    if _is_email_pdf_spv:
        invoice_mode = "AUTO-GENERATE (SPV rule; PDF extraction disabled)"
    elif provided_invoice:
        invoice_mode = "USE PROVIDED INVOICE NUMBER"
    elif skip_extraction:
        invoice_mode = "AUTO-GENERATE (PDF extraction disabled globally)"
    elif vendor_code_norm in config.NIS_AUTO_INVOICE_VENDOR_CODES:
        invoice_mode = f"AUTO-GENERATE (vendor rule: {vendor_code_norm})"
    else:
        invoice_mode = "EXTRACT FROM PDF (auto-generate only if extraction fails)"

    print(f"  [INVOICE-DECISION] SPV: {_spv_name_for_check or '(missing)'}")
    print(f"  [INVOICE-DECISION] Mode: {invoice_mode}")
    lm.log(
        f"NIS_INVOICE_DECISION spv={_spv_name_for_check or '(missing)'} "
        f"vendor={vendor_code_norm or '(missing)'} mode={invoice_mode}"
    )
    # ────────────────────────────────────────────────────────────────────────

    if provided_invoice:

        # Highest priority: an invoice number resolved by the caller (e.g. the
        # 3-row rule, where each row's invoice number is pre-extracted from the
        # attached PDF — per-page for an INOX bill, or one shared number when
        # the bill has a single page).
        invoice_no = provided_invoice
        lm.log(f"NIS_INVOICE_PROVIDED vendor={vendor_code_norm} invoice={invoice_no}")
        print(f"✅ Using caller-provided invoice number: {invoice_no}")
    elif skip_extraction or vendor_code_norm in config.NIS_AUTO_INVOICE_VENDOR_CODES:
        invoice_no = generated_invoice_no
        reason = "extraction disabled" if skip_extraction else f"vendor {vendor_code_norm}"
        lm.log(f"NIS_INVOICE_AUTO_GENERATED vendor={vendor_code_norm} invoice={invoice_no} ({reason})")
        print(f"✅ Auto-generated invoice number ({reason}): {invoice_no}")
    else:
        invoice_no = ""
        try:
            from invoice_extractor import extract_invoice_number_from_pdf
            invoice_no = extract_invoice_number_from_pdf(pdf_path)
        except Exception as exc:
            lm.log_warn(f"NIS_INVOICE_EXTRACT_FAILED vendor={vendor_code_norm}: {exc}")
            invoice_no = ""
        if invoice_no:
            lm.log(f"NIS_INVOICE_EXTRACTED vendor={vendor_code_norm} invoice={invoice_no}")
            print(f"✅ Vendor {vendor_code_norm} → invoice number extracted from PDF: {invoice_no}")
        else:
            # Could not read an invoice number from the PDF — fall back to the
            # auto-generated value so the booking can still proceed.
            invoice_no = generated_invoice_no
            lm.log_warn(
                f"NIS_INVOICE_EXTRACT_EMPTY vendor={vendor_code_norm} "
                f"falling_back_to_generated={invoice_no}"
            )
            print(
                f"⚠️ Vendor {vendor_code_norm} → could not extract invoice number "
                f"from PDF, using generated value: {invoice_no}"
            )

    lm.log(f"NIS_FILL_VENDOR vendor={vendor_code} bank={bank_key} "
           f"invoice={invoice_no} date={invoice_date}")
    try:
        fill_vendor_field(frame, vendor_code)
        fill_bank_key_field(frame, bank_key)
        fill_invoice_number_field(frame, invoice_no)
        fill_invoice_date_field(frame, invoice_date)
        lm.log("NIS_FILL_VENDOR_OK")
    except Exception as exc:
        lm.log_error("NIS_FILL_VENDOR_FAILED", exc=exc)
        raise

    try:
        ensure_switch_on(page, frame)
        lm.log("NIS_SWITCH_ON")
    except Exception as exc:
        lm.log_error("NIS_SWITCH_ON_FAILED", exc=exc)
        raise

    try:
        fill_payment_method(page, frame)
        lm.log("NIS_FILL_PAYMENT_METHOD_OK")
    except Exception as exc:
        lm.log_error("NIS_FILL_PAYMENT_METHOD_FAILED", exc=exc)
        raise
    time.sleep(2)

    # Go to Add Item screen
    try:
        lm.log("NIS_CLICK_NEXT_TO_ADD_ITEMS")
        click_next_button3(frame)
        lm.log("NIS_CLICK_NEXT_TO_ADD_ITEMS_OK")
    except Exception as exc:
        lm.log_error("NIS_CLICK_NEXT_TO_ADD_ITEMS_FAILED", exc=exc)
        raise

    # ── Wait for the Items tab content to fully load ─────────────────────────
    # After clicking Next, SAP UI5 transitions to the Items tab.  We must NOT
    # proceed until the "Add Expense Item" button is actually in the DOM and
    # visible.  A fixed sleep is unreliable on slow VMs — instead we:
    #   1) Wait for the Items tab to become the selected tab (nav done).
    #   2) Wait for any loading / busy indicators to clear.
    #   3) Wait for the Add Expense Item button to be visible (up to 60 s).
    try:
        # 1) Items tab selected — confirms the wizard navigated to step 2
        items_tab = frame.get_by_role("tab", name=re.compile(r"Items", re.I))
        try:
            expect(items_tab.first).to_have_attribute("aria-selected", "true", timeout=20000)
        except Exception:
            # Tab may use 'selected' class instead — just wait a moment
            time.sleep(3)
    except Exception:
        time.sleep(3)

    # 2) Wait for loading spinners / block layers to clear
    for _sel in [".sapUiBlockLayer", ".sapUiLocalBusyIndicator", ".sapMBusyDialog"]:
        try:
            frame.locator(_sel).wait_for(state="detached", timeout=10000)
        except Exception:
            pass

    # 3) Wait for the Add Expense Item button to be visible (up to 60 s)
    try:
        _add_btn = frame.locator(
            "#application-Nis-manage-component---CreateNIS--AddItemsButton"
        )
        if _add_btn.count() == 0:
            _add_btn = frame.get_by_role("button", name=re.compile(r"Add Expense Item", re.I))
        _add_btn.first.wait_for(state="visible", timeout=60000)
    except Exception:
        # If still not visible after 60 s, log and let click_add_expense_item raise
        lm.log_warn("NIS_ADD_ITEMS_BTN_NOT_VISIBLE_AFTER_WAIT")
    # ─────────────────────────────────────────────────────────────────────────


    # -------- Add Item section based on SVP DSM table ---------

    first_item = True  # Track whether this is the first item or a subsequent one

    def add_one_item(amount, label, item_cost_center=None, item_plant=None,
                     item_purpose=None, item_additional="PSS8", field_key=None):
        nonlocal first_item

        if amount is None:
            return
        try:
            amount_val = float(amount)
        except Exception:
            return
        if amount_val <= 0:
            return

        # Per-item vendor fields default to the booking-level ones. For a
        # combined Wind/Solar booking each variant supplies its own cost
        # center / plant so the items book against the correct master data.
        icc = item_cost_center if item_cost_center not in (None, "") else cost_center
        ipl = item_plant if item_plant not in (None, "") else plant
        ipu = item_purpose if item_purpose not in (None, "") else purpose

        # Resolve the correct Expense Description (sub-category) for this charge type.
        subcategory_value = EXPENSE_SUBCATEGORY_MAP.get(
            field_key, _DEFAULT_EXPENSE_SUBCATEGORY
        ) if field_key else _DEFAULT_EXPENSE_SUBCATEGORY

        print(f"Adding item for {label}: {amount_val}")
        lm.log(f"NIS_ADD_ITEM label={label} amount={amount_val} "
               f"cost_center={icc} plant={ipl} subcategory={subcategory_value}")

        # FIX: For the first item, the "Add Expense Item" dialog is opened by clicking
        # the Add button.  For subsequent items, we must first wait for the previous
        # dialog's DOM (especially the WBS input that lives in #sap-ui-static) to
        # fully detach before clicking the Add button again, otherwise the WBS input
        # intercepts the pointer event and causes a 30 s timeout.
        if not first_item:
            _dismiss_sap_ui_static_overlay(page, frame)
        else:
            first_item = False

        # Click "Add Expense Item" — now uses the updated signature with page + frame
        # and includes its own WBS-detach guard as a safety net.
        click_add_expense_item(page, frame)
        time.sleep(2)

        # Category & sub-category
        type_expense_category_label(page, frame)
        time.sleep(2)
        type_expense_subcategory(page, frame, value=subcategory_value)
        time.sleep(2)

        # Amount
        fill_amount_value(frame, amount=str(amount_val))
        time.sleep(1)

        # Cost center, plant & purpose
        fill_cost_center(frame, value=icc)
        fill_plant(frame, value=ipl)
        fill_remarks(frame, value=ipu)
        fill_additional_remarks(frame, value=item_additional)

        # Tick the acknowledge checkbox
        click_acknowledge_checkbox(page, frame)

        # Confirm item
        click_ok_button(page, frame)
        lm.log(f"NIS_ADD_ITEM_OK label={label} amount={amount_val}")
        time.sleep(2)

        # ── Optional splash / info popup after Add Item ────────────────
        # The portal may show a MessageBox with an OK button (id=__mbox-btn-0)
        # after each item is added.  It may or may not appear — handle both.
        _dismiss_add_item_popup(page, frame)

    # ── Build the expense items ──────────────────────────────────────────────
    # When a Wind/Solar pair is booked together, main.py passes an explicit
    # `expense_items` list (each with its own cost center / plant). Otherwise we
    # derive the items from this single SPV row's positive amount fields.
    if expense_items:
        lm.log(f"NIS_ADD_ITEMS_START combined_items={len(expense_items)}")
        for it in expense_items:
            add_one_item(
                it.get("amount"),
                it.get("label", "Expense item"),
                item_cost_center=it.get("cost_center"),
                item_plant=it.get("plant"),
                item_purpose=it.get("purpose"),
                item_additional=it.get("additional_remarks", "PSS8"),
                field_key=it.get("field_key"),
            )
    else:
        lm.log(f"NIS_ADD_ITEMS_START row={svp_dsm_row}")
        add_one_item(svp_dsm_row.get("total_dsm_charges_payable"), "Total DSM charges payable",
                     field_key="total_dsm_charges_payable")
        add_one_item(svp_dsm_row.get("drawl_charges_payable"), "Drawl charges payable",
                     field_key="drawl_charges_payable")
        add_one_item(svp_dsm_row.get("revenue_diff"), "Revenue difference",
                     field_key="revenue_diff")
        add_one_item(svp_dsm_row.get("revenue_loss"), "Revenue loss",
                     field_key="revenue_loss")
    lm.log("NIS_ADD_ITEMS_DONE")

    # After all items, proceed to approval screen
    _dismiss_sap_ui_static_overlay(page, frame)
    try:
        lm.log("NIS_CLICK_NEXT_TO_APPROVAL")
        click_next(frame)
        lm.log("NIS_CLICK_NEXT_TO_APPROVAL_OK")
    except Exception as exc:
        lm.log_error("NIS_CLICK_NEXT_TO_APPROVAL_FAILED", exc=exc)
        raise
    time.sleep(3)

    # -------- Approval Final section ---------
    try:
        lm.log("NIS_CLICK_FINAL_NEXT_1")
        _click_final_next(page, frame)
        lm.log("NIS_CLICK_FINAL_NEXT_1_OK")
    except Exception as exc:
        lm.log_error("NIS_CLICK_FINAL_NEXT_1_FAILED", exc=exc)
        raise

    # -------- Select approver radio in fcTable ---------
    # The radio button is NOT always present (depends on approval workflow
    # configuration). If not found within timeout, skip it and proceed
    # directly to Submit.
    page.wait_for_timeout(2000)
    radio_found = False
    try:
        lm.log("NIS_SELECT_APPROVER_RADIO")
        _click_fcTable_radio(page, frame)
        lm.log("NIS_SELECT_APPROVER_RADIO_OK")
        radio_found = True
    except Exception as exc:
        lm.log_warn(f"NIS_SELECT_APPROVER_RADIO_SKIPPED (not found — proceeding to submit): {exc}")
        print(f"  ⚠️ Approval radio not found — skipping (may not be required for this booking).")

    if radio_found:
        page.wait_for_timeout(2000)
        try:
            lm.log("NIS_CLICK_FINAL_NEXT_2")
            _click_final_next(page, frame)
            lm.log("NIS_CLICK_FINAL_NEXT_2_OK")
        except Exception as exc:
            lm.log_warn(f"NIS_CLICK_FINAL_NEXT_2_SKIPPED: {exc}")
            print(f"  ⚠️ Second 'Next' after radio not found — proceeding to submit.")

    # -------- Submit the NIS booking --------
    page.wait_for_timeout(3000)
    try:
        lm.log("NIS_SUBMIT")
        click_submit_button(page, frame)
        lm.log("NIS_SUBMIT_OK")
    except Exception as exc:
        lm.log_error("NIS_SUBMIT_FAILED", exc=exc)
        raise

    # -------- Capture NIS checklist ID from success popup --------
    checklist_id = ""
    page.wait_for_timeout(5000)
    try:
        lm.log("NIS_CAPTURE_CHECKLIST_ID")
        checklist_id = capture_nis_checklist_id(page, frame)
        if checklist_id:
            lm.log(f"NIS_CHECKLIST_ID_CAPTURED id={checklist_id}")
            print(f"✅ NIS Checklist ID captured: {checklist_id}")
        else:
            lm.log_warn("NIS_CHECKLIST_ID_NOT_CAPTURED")
            print("⚠️  NIS Checklist ID could not be captured from the screen.")
    except Exception as exc:
        lm.log_error("NIS_CAPTURE_CHECKLIST_ID_FAILED", exc=exc)
        # non-fatal — still return what we have

    return checklist_id


def capture_nis_checklist_id(page, frame, timeout: int = 30000) -> str:
    """
    After Submit, the NIS portal shows a success MessageBox dialog:

        <span class="sapMMsgBoxText">Request 7000088592 submitted successfully.</span>

    The dialog typically renders in the ``#sap-ui-static`` overlay layer
    (at the top-level page), NOT inside the NIS iframe.  Its auto-generated
    id pattern is ``__success0`` with a footer ``__success0-footer``.

    This function:
      1. Waits for the success dialog to appear (polls across multiple scopes).
      2. Extracts the NIS booking number (the numeric part) from the message text.
      3. Clicks the OK button (``#__mbox-btn-1``, ``data-ui5-accesskey='o'``) to
         dismiss the popup.
      4. Returns the NIS booking number as a string, or "" on failure.
    """

    import time as _time

    nis_number = ""

    iframe_selector = (
        "iframe[id^='application-Nis-manage-'], "
        "iframe[name^='application-Nis-manage-'], "
        "iframe[src*='Nis-manage']"
    )

    # ── Helper: wait for the success dialog to be visible ────────────────────
    def _wait_for_success_dialog() -> bool:
        """
        Poll the page until the success dialog's message text is visible.
        Returns True when found, False on timeout.
        """
        deadline = _time.time() + timeout / 1000
        selectors = [
            ".sapMMsgBoxText",
            "span:has-text('submitted successfully')",
            "[id^='__success'] .sapMText",
        ]
        while _time.time() < deadline:
            # Check #sap-ui-static first (most likely location)
            for sel in selectors:
                try:
                    loc = page.locator(f"#sap-ui-static {sel}")
                    if loc.count() > 0 and loc.first.is_visible():
                        return True
                except Exception:
                    pass

            # Check top-level page
            for sel in selectors:
                try:
                    loc = page.locator(sel)
                    if loc.count() > 0 and loc.first.is_visible():
                        return True
                except Exception:
                    pass

            page.wait_for_timeout(500)
        return False

    # ── Step 0: Wait for the popup to actually appear ────────────────────────
    print("[NIS ID] Waiting for success popup to appear …")
    if not _wait_for_success_dialog():
        print("[NIS ID] ⚠️  Success popup did not appear within timeout.")
        # Still attempt extraction below — it may be present but our check missed it.

    # ── Helper: extract NIS number from a scope + dismiss OK ─────────────────
    def _extract_and_dismiss(scope, label: str) -> str:
        nonlocal nis_number

        # --- Find the success message text ---
        msg_locators = [
            scope.locator(".sapMMsgBoxText"),
            scope.locator("[id^='__success'] .sapMText"),
            scope.locator("[role='dialog'] span.sapMText, [role='alertdialog'] span.sapMText"),
            scope.locator("span:has-text('submitted successfully')"),
        ]

        msg_text = ""
        for loc in msg_locators:
            try:
                if loc.count() == 0:
                    continue
                loc.first.wait_for(state="visible", timeout=5000)
                msg_text = loc.first.inner_text().strip()
                if msg_text:
                    break
            except Exception:
                continue

        if msg_text:
            print(f"[NIS ID] Success message found [{label}]: {msg_text}")
            match = re.search(r'\b(\d{7,})\b', msg_text)
            if match:
                nis_number = match.group(1)
                print(f"✅ NIS Booking Number captured: {nis_number}")
            else:
                print(f"⚠️  Could not parse NIS number from message: {msg_text}")
        else:
            print(f"[NIS ID] No success message text found in [{label}]")
            return ""

        # --- Click the OK button to dismiss the popup ---
        ok_strategies = [
            # 1) Exact id of the OK button seen in the footer HTML
            lambda s: s.locator("#__mbox-btn-1"),
            # 2) Any __mbox-btn with OK text
            lambda s: s.locator("[id^='__mbox-btn']").filter(has_text=re.compile(r"^\s*OK\s*$", re.I)),
            # 3) SAP accesskey="o" (the OK button carries this attribute)
            lambda s: s.locator("button[data-ui5-accesskey='o']"),
            # 4) Emphasized button with OK text inside <bdi>
            lambda s: s.locator("button.sapMBtnEmphasized:has(bdi:text-is('OK'))"),
            # 5) ARIA role + name
            lambda s: s.get_by_role("button", name=re.compile(r"^\s*OK\s*$", re.I)),
            # 6) Generic has-text fallback (restrict to dialog footer)
            lambda s: s.locator(".sapMDialogFooter button:has-text('OK')"),
            # 7) Broadest fallback
            lambda s: s.locator("button:has-text('OK')"),
        ]

        for strategy in ok_strategies:
            try:
                btn = strategy(scope)
                if btn.count() == 0:
                    continue
                btn.first.wait_for(state="visible", timeout=5000)
                try:
                    btn.first.scroll_into_view_if_needed()
                except Exception:
                    pass
                try:
                    btn.first.click(timeout=5000)
                    print(f"✅ OK button clicked — success popup dismissed [{label}]")
                    return nis_number
                except Exception:
                    try:
                        btn.first.evaluate("el => el.click()")
                        print(f"✅ OK button clicked via JS — success popup dismissed [{label}]")
                        return nis_number
                    except Exception:
                        pass
            except Exception:
                continue

        print(f"⚠️  OK button not found in [{label}] — trying JS fallback")
        return nis_number

    # ── Scope 1: #sap-ui-static overlay (SAP places dialogs here) ───────────
    try:
        static_layer = page.locator("#sap-ui-static")
        if static_layer.count() > 0:
            result = _extract_and_dismiss(static_layer, "#sap-ui-static")
            if result:
                return result
    except Exception:
        pass

    # ── Scope 2: top-level page ──────────────────────────────────────────────
    result = _extract_and_dismiss(page, "top-level page")
    if result:
        return result

    # ── Scope 3: existing frame locator ──────────────────────────────────────
    result = _extract_and_dismiss(frame, "existing frame locator")
    if result:
        return result

    # ── Scope 4: fresh frame_locator ─────────────────────────────────────────
    try:
        fresh_frame = page.frame_locator(iframe_selector)
        result = _extract_and_dismiss(fresh_frame, "fresh frame_locator")
        if result:
            return result
    except Exception:
        pass

    # ── Scope 5: JS walk across all frames (last resort) ────────────────────
    try:
        js_result = page.evaluate("""
            () => {
                // Check #sap-ui-static first
                const staticEl = document.getElementById('sap-ui-static');
                const roots = staticEl ? [staticEl, document] : [document];

                for (const root of roots) {
                    // Find the message text
                    const msgEl = root.querySelector('.sapMMsgBoxText')
                               || root.querySelector('[id^="__success"] .sapMText')
                               || root.querySelector('[role="dialog"] span.sapMText');
                    if (!msgEl) continue;
                    const text = msgEl.textContent.trim();
                    const match = text.match(/\\b(\\d{7,})\\b/);
                    const nisNum = match ? match[1] : '';

                    // Click OK — try several selectors
                    const okSelectors = [
                        '#__mbox-btn-1',
                        "[id^='__mbox-btn']",
                        "button[data-ui5-accesskey='o']",
                        '.sapMDialogFooter button.sapMBtnEmphasized',
                    ];
                    for (const sel of okSelectors) {
                        const btn = root.querySelector(sel);
                        if (btn && btn.textContent.trim().toUpperCase().includes('OK')) {
                            btn.click();
                            break;
                        }
                    }
                    return nisNum;
                }

                // Also try inside iframes
                const frames = Array.from(window.frames);
                for (const f of frames) {
                    try {
                        const msgEl = f.document.querySelector('.sapMMsgBoxText')
                                   || f.document.querySelector('[role="dialog"] span.sapMText');
                        if (!msgEl) continue;
                        const text = msgEl.textContent.trim();
                        const match = text.match(/\\b(\\d{7,})\\b/);
                        const nisNum = match ? match[1] : '';

                        const okBtns = f.document.querySelectorAll(
                            "#__mbox-btn-1, button[data-ui5-accesskey='o'], [id^='__mbox-btn']"
                        );
                        for (const btn of okBtns) {
                            if (btn.textContent.trim().toUpperCase().includes('OK')) {
                                btn.click();
                                break;
                            }
                        }
                        return nisNum;
                    } catch (e) {}
                }
                return '';
            }
        """)
        if js_result:
            nis_number = js_result
            print(f"✅ NIS Booking Number captured via JS walk: {nis_number}")
            return nis_number
    except Exception:
        pass

    if nis_number:
        return nis_number

    print("⚠️  NIS booking number could not be captured from the success popup.")
    return ""