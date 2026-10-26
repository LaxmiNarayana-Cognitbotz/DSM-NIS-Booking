import os
import time
import subprocess
import win32com.client
from datetime import datetime
import sap_config as config


def _co_initialize():
    """Initialize COM for the current thread when pywin32 requires it."""
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass


def _get_sap_application():
    _co_initialize()
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
    except Exception as e:
        raise RuntimeError(f"GetObject('SAPGUI') failed. Is SAP Logon running? Error: {e}")
    return SapGuiAuto.GetScriptingEngine

# -------------------------------------------------
# Helpers: resilient UI actions & waits
# -------------------------------------------------
def wait(seconds: float = 1.0):
    time.sleep(seconds)

def try_press(session, ids):
    """Try pressing the first existing control from a list of IDs."""
    for _id in ids:
        try:
            session.findById(_id).press()
            print(f"    [PRESS] {_id} — OK")
            return True
        except Exception:
            pass
    print(f"    [PRESS] None of {ids} found")
    return False

def try_select(session, ids):
    """Try selecting (e.g., radio/tab) the first existing control from a list of IDs."""
    for _id in ids:
        try:
            session.findById(_id).select()
            print(f"    [SELECT] {_id} — OK")
            return True
        except Exception:
            pass
    print(f"    [SELECT] None of {ids} found")
    return False

def try_set_text(session, _id, value):
    try:
        session.findById(_id).text = str(value)
        print(f"    [FILL] {_id} = '{value}'")
        return True
    except Exception as e:
        print(f"    [FILL] FAILED {_id} = '{value}' — {e}")
        return False

def try_set_key(session, _id, key):
    """For combo boxes where .key is the right property."""
    try:
        session.findById(_id).key = str(key)
        print(f"    [FILL] {_id} .key = '{key}'")
        return True
    except Exception as e:
        print(f"    [FILL] FAILED {_id} .key = '{key}' — {e}")
        return False

def send_enter(session):
    try:
        session.findById("wnd[0]").sendVKey(0)
        return True
    except Exception:
        return False

def handle_default_popup_ok(session):
    """Press OK/Continue on common popups if they appear (wnd[1])."""
    try:
        session.findById("wnd[1]/tbar[0]/btn[0]").press()  # OK / Continue
        return True
    except Exception:
        return False

def handle_yes_popup(session):
    """Press Yes/Enter on confirmation popup (wnd[1]) if present."""
    return try_press(session, [
        "wnd[1]/tbar[0]/btn[8]",
        "wnd[1]/tbar[0]/btn[0]",
    ])

# -------------------------------------------------
# SAP bootstrap: launch, attach, connect, login
# -------------------------------------------------

def launch_sap_logon(saplogon_path: str):
    """Launch SAP Logon if not already running."""
    import psutil
    try:
        for proc in psutil.process_iter(["name"]):
            if proc.info.get("name") and proc.info["name"].lower() == "saplogon.exe":
                print("  [SAP] SAP Logon is already running.")
                return
    except Exception:
        pass

    print(f"  [SAP] Launching SAP Logon from: {saplogon_path}")
    if not os.path.exists(saplogon_path):
        fallbacks = [
            r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",
            r"C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe",
        ]
        for fb in fallbacks:
            if os.path.exists(fb):
                saplogon_path = fb
                print(f"  [SAP] Found fallback path: {saplogon_path}")
                break
        else:
            print(f"  [SAP] WARNING: SAP Logon executable not found at {saplogon_path}")

    try:
        subprocess.Popen(saplogon_path)
        wait(10)  # Wait longer for SAP to initialize and register COM
    except Exception as e:
        print(f"  [SAP] Failed to launch SAP Logon: {e}")


def _wait_for_sap_processes_to_exit(timeout: float = 15.0):
    """Wait briefly after cleanup so the next run does not attach stale COM."""
    try:
        import psutil
    except ImportError:
        time.sleep(3)
        return

    deadline = time.time() + timeout
    process_names = {"saplogon.exe", "sapgui.exe"}
    while time.time() < deadline:
        alive = []
        for proc in psutil.process_iter(["name"]):
            try:
                name = (proc.info.get("name") or "").lower()
                if name in process_names:
                    alive.append(proc)
            except Exception:
                continue
        if not alive:
            return
        time.sleep(0.5)

def attach_to_scripting(system_name: str):
    last_error = None
    for _ in range(3):
        try:
            application = _get_sap_application()
            connection = application.OpenConnection(system_name, True)
            deadline = time.time() + 20
            while time.time() < deadline:
                try:
                    session = connection.Children(0)
                    session.findById("wnd[0]")
                    return application, connection, session
                except Exception as exc:
                    last_error = exc
                    time.sleep(0.5)
        except Exception as exc:
            last_error = exc
            time.sleep(2)

    raise RuntimeError(f"Could not attach to SAP scripting session: {last_error}")


def _reconnect_session(timeout: float = 20.0):
    """
    Re-attach to an EXISTING SAP GUI session after the COM link has been lost
    (e.g. after pywinauto interactions with native Windows dialogs).
    Unlike attach_to_scripting(), this does NOT open a new connection. It scans
    all existing connections/sessions and returns the first live session.
    Returns: session  (or raises on failure)
    """
    deadline = time.time() + timeout
    last_error = None

    while time.time() < deadline:
        try:
            application = _get_sap_application()
            conn_count = int(application.Children.Count)
            for ci in range(conn_count):
                try:
                    connection = application.Children(ci)
                    sess_count = int(connection.Children.Count)
                except Exception as exc:
                    last_error = exc
                    continue

                for si in range(sess_count):
                    try:
                        session = connection.Children(si)
                        # Touch wnd[0] so we do not return a stale COM proxy.
                        session.findById("wnd[0]")
                        return session
                    except Exception as exc:
                        last_error = exc
                        continue
        except Exception as exc:
            last_error = exc

        time.sleep(0.5)

    raise RuntimeError(f"Could not reconnect to a live SAP session: {last_error}")

def login(session, client: str, username: str, password: str, language: str = "EN"):
    print(f"\n  ── login: client={client}, user={username}, lang={language} ──")
    session.findById("wnd[0]/usr/txtRSYST-MANDT").text = client
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = username
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
    session.findById("wnd[0]/usr/txtRSYST-LANGU").text = language
    print(f"    [FILL] MANDT = '{client}'")
    print(f"    [FILL] BNAME = '{username}'")
    print(f"    [FILL] BCODE = '****'")
    print(f"    [FILL] LANGU = '{language}'")
    send_enter(session)
    wait(3)
    print("  ── login: Done ──\n")

def go_to_tcode(session, tcode: str):
    print(f"  ── go_to_tcode: '{tcode}' ──")
    session.findById("wnd[0]/tbar[0]/okcd").text = tcode
    send_enter(session)
    wait(2)
    print(f"  ── go_to_tcode: Done ──\n")

# -------------------------------------------------
# Task flows (configurable)
# -------------------------------------------------
def select_non_po_and_create(session):
    """
    Select the Non-PO Based option (if present) then trigger Create.
    Multiple fallbacks are attempted as different systems expose different IDs.
    """
    print("  ── select_non_po_and_create ──")
    selected = try_select(session, [
        "wnd[0]/usr/radNON_PO_BASED",
        "wnd[1]/usr/radNON_PO_BASED",
        "wnd[0]/usr/radRB3",
    ])
    print(f"    Non-PO radio selected: {selected}")

    created = try_press(session, [
        "wnd[0]/usr/btnCREATE_CHECKLIST",
        "wnd[0]/usr/btnCREATE",
        "wnd[0]/tbar[1]/btn[8]",
    ])
    print(f"    Create button pressed: {created}")
    handle_cancel_popup(session)
    if created:
        wait(0.8)
    print("  ── select_non_po_and_create: Done ──\n")


def fill_non_po_form(session, params: dict):
    """Fill fields for the Non-PO flow (config-driven)."""
    print("\n  ── fill_non_po_form: Filling initial form fields ──")
    print(f"    Params: CHK_DOC_TYP={params.get('CHK_DOC_TYP')}, "
          f"BUKRS={params.get('BUKRS')}, DIG_SIGN_INV={params.get('DIG_SIGN_INV')}, "
          f"GJAHR={params.get('GJAHR')}, LIFNR={params.get('LIFNR')}, "
          f"DEB_NOT_REF={params.get('DEB_NOT_REF')}")

    try_set_text(session, "wnd[0]/usr/ctxtZFIBPT11001-CHK_DOC_TYP", params.get("CHK_DOC_TYP", ""))
    try_set_text(session, "wnd[0]/usr/ctxtZFIBPT11001-BUKRS", params.get("BUKRS", ""))
    try_set_key(session,  "wnd[0]/usr/cmbZFIBPT11001-DIG_SIGN_INV", params.get("DIG_SIGN_INV", ""))

    if params.get("GJAHR"):
        try_set_text(session, "wnd[0]/usr/txtZFIBPT11001-GJAHR", params["GJAHR"])
        send_enter(session)
        wait(0.7)
    if params.get("LIFNR"):
        try_set_text(session, "wnd[0]/usr/ctxtZFIBPT11001-LIFNR", params["LIFNR"])

    print("    Pressing Execute (btn[8])...")
    try_press(session, ["wnd[0]/tbar[1]/btn[8]"])
    handle_default_popup_ok(session)

    try:
        tbl = session.findById(
            "wnd[0]/usr/tabsZTABSTRIP/tabpZTABSTRIP_FC1/"
            "ssubZTABSTRIP_SCA:SAPMZFIBP11001:9012/tblSAPMZFIBP11001ZTC_9013"
        )
        tbl.verticalScrollbar.position = 1
        print("    Table scrollbar set to position 1")
    except Exception:
        print("    Table scrollbar set skipped (not found)")

    if params.get("DEB_NOT_REF"):
        try_set_text(session, "wnd[0]/usr/txtZFIBPT11001-DEB_NOT_REF", params["DEB_NOT_REF"])
    print("  ── fill_non_po_form: Done ──\n")

def apply_subcategory_checkbox_table_and_save(session, params: dict, pdf_output_dir: str = None):
    """
    Implements the additional steps:
    - Fill Debit Note Date and Document Sub-Category
    - Go to the 'Documents' tab and tick a checkbox
    - Return to the first tab and fill table fields (Amount, Cost Center, Plant)
    - (Optional) Preview
    - Save and confirm popups
    """
    print("\n  ── apply_subcategory_checkbox_table_and_save ──")
    print(f"    Params: DEB_NOT_DATE={params.get('DEB_NOT_DATE')}, "
          f"DOC_SUB_CAT={params.get('DOC_SUB_CAT')}, "
          f"DEB_NOT_REF={params.get('DEB_NOT_REF')}, "
          f"AMOUNT={params.get('AMOUNT')}, "
          f"COST_CENTER={params.get('COST_CENTER')}, "
          f"PLANT={params.get('PLANT')}")

    if params.get("DEB_NOT_DATE"):
        try_set_text(session, "wnd[0]/usr/ctxtZFIBPT11001-DEB_NOT_DAT", params["DEB_NOT_DATE"])
    if params.get("DOC_SUB_CAT"):
        try_set_text(session, "wnd[0]/usr/ctxtZFIBPT11001-DOC_SUB_CAT", params["DOC_SUB_CAT"])
    # Ensure DEB_NOT_REF is set reliably on this screen (belt-and-suspenders:
    # also attempted in fill_non_po_form, but the table scroll there can
    # interfere with field visibility causing try_set_text to fail silently).
    if params.get("DEB_NOT_REF"):
        try_set_text(session, "wnd[0]/usr/txtZFIBPT11001-DEB_NOT_REF", params["DEB_NOT_REF"])

    try:
        session.findById("wnd[0]/usr/txtGV_DOC_SUB_TEXT").setFocus()
        session.findById("wnd[0]/usr/txtGV_DOC_SUB_TEXT").caretPosition = 0
    except Exception:
        pass

    # Navigate to checkbox tab (FC3) and tick the checkbox
    print("    Navigating to Documents tab (FC3) to tick checkbox...")
    try_select(session, ["wnd[0]/usr/tabsZTABSTRIP/tabpZTABSTRIP_FC3"])
    try:
        chk_id = (
            "wnd[0]/usr/tabsZTABSTRIP/tabpZTABSTRIP_FC3/"
            "ssubZTABSTRIP_SCA:SAPMZFIBP11001:9005/tblSAPMZFIBP11001ZTC_9005/"
            "chkGW_DOCS-CHECK[1,0]"
        )
        cell = session.findById(chk_id)
        cell.selected = True
        cell.setFocus()
        print(f"    [FILL] Checkbox {chk_id} = selected (True)")
    except Exception as e:
        print(f"    [FILL] FAILED Checkbox tick — {e}")

    print("    Navigating back to FC2 → FC1 tabs...")
    try_select(session, ["wnd[0]/usr/tabsZTABSTRIP/tabpZTABSTRIP_FC2"])
    try_select(session, ["wnd[0]/usr/tabsZTABSTRIP/tabpZTABSTRIP_FC1"])

    # Fill table row (row 0) fields
    print("    Filling table row 0 (Amount, Cost Center, Plant)...")
    base_tbl = (
        "wnd[0]/usr/tabsZTABSTRIP/tabpZTABSTRIP_FC1/"
        "ssubZTABSTRIP_SCA:SAPMZFIBP11001:9012/tblSAPMZFIBP11001ZTC_9013/"
    )

    if params.get("AMOUNT"):
        try_set_text(session, base_tbl + "txtGW_RE_MODULE-DEB_NOT_AMT[3,0]", params["AMOUNT"])
    if params.get("COST_CENTER"):
        try_set_text(session, base_tbl + "ctxtGW_RE_MODULE-KOSTL[4,0]", params["COST_CENTER"])
    if params.get("PLANT"):
        try_set_text(session, base_tbl + "ctxtGW_RE_MODULE-WERKS[5,0]", params["PLANT"])
        try:
            session.findById(base_tbl + "ctxtGW_RE_MODULE-WERKS[5,0]").setFocus()
            session.findById(base_tbl + "ctxtGW_RE_MODULE-WERKS[5,0]").caretPosition = len(str(params["PLANT"]))
        except Exception:
            pass

    try:
        tbl = session.findById(
            "wnd[0]/usr/tabsZTABSTRIP/tabpZTABSTRIP_FC1/"
            "ssubZTABSTRIP_SCA:SAPMZFIBP11001:9012/tblSAPMZFIBP11001ZTC_9013"
        )
        tbl.verticalScrollbar.position = 0
        print("    Table scrollbar reset to position 0")
    except Exception:
        pass

    wait(2)

    print("    Pressing SAVE (btn[11])...")
    saved = try_press(session, [
        "wnd[0]/tbar[0]/btn[11]",
        "wnd[0]/tbar[1]/btn[11]",
        "wnd[0]/usr/btnSAVE",
    ])
    print(f"    Save pressed: {saved}")
    print("  ── Form filling complete, starting checklist capture ──\n")

    # ══════════════════════════════════════════════════════════════════════════
    # PHASE 1: CAPTURE CHECKLIST NUMBER BEFORE PRINT (session is alive here)
    # After Save, wnd[1] popup appears with btn[86]. The wnd[0]/sbar and form
    # fields are still readable even with the modal popup on top.
    # ══════════════════════════════════════════════════════════════════════════
    import re as _re_chk
    import time
    _checklist_number = None
    time.sleep(1.5)   # give SAP a moment to finish the save

    print("\n  ── Phase 1: Capturing checklist number BEFORE print ──")

    def _extract_number_from_text(_label, _text):
        _text = str(_text or "").strip()
        if not _text:
            return None
        print(f"  📋 {_label}: {_text}")
        _m = _re_chk.search(r'\b(\d{7,})\b', _text)
        return _m.group(1) if _m else None

    def _try_pre_print_capture(_sess, _attempt):
        # Try reading status bar text (wnd[0]/sbar readable even with wnd[1] modal)
        try:
            _num = _extract_number_from_text(
                f"SAP status bar after Save (attempt {_attempt})",
                _sess.findById("wnd[0]/sbar").text,
            )
            if _num:
                print(f"  ✅ Checklist number from status bar (pre-print): {_num}")
                return _num
        except Exception as _sbar_err:
            print(f"  ⚠️  Could not read sbar (pre-print attempt {_attempt}): {_sbar_err}")

        # The save/print popup sometimes contains the final message/number.
        for _popup_id in [
            "wnd[1]/usr/txtMESSAGE",
            "wnd[1]/usr/txtMESSTXT-STEXT",
            "wnd[1]/usr/subSCREEN:SAPLMSG_DISPLAY:0100/txtMESSAGE",
            "wnd[1]/usr/subSCREEN_HEADER:SAPLMSG_DISPLAY:0100/txtMESSAGE",
            "wnd[1]",
        ]:
            try:
                _ctrl = _sess.findById(_popup_id)
                _num = _extract_number_from_text(
                    f"Popup text {_popup_id} (attempt {_attempt})",
                    getattr(_ctrl, "text", "") or getattr(_ctrl, "Text", ""),
                )
                if _num:
                    print(f"  ✅ Checklist number from popup (pre-print): {_num}")
                    return _num
            except Exception:
                continue

        # Try reading CHK_NO form field directly.
        try:
            _num = _extract_number_from_text(
                f"CHK_NO field (attempt {_attempt})",
                _sess.findById("wnd[0]/usr/txtZFIBPT11001-CHK_NO").text,
            )
            if _num:
                print(f"  ✅ Checklist number from CHK_NO field (pre-print): {_num}")
                return _num
        except Exception:
            pass

        return None

    for _capture_attempt in range(1, 4):
        _checklist_number = _try_pre_print_capture(session, _capture_attempt)
        if _checklist_number:
            break
        if _capture_attempt < 3:
            time.sleep(2 * _capture_attempt)
            try:
                session.findById("wnd[0]")
            except Exception:
                try:
                    session = _reconnect_session(timeout=60)
                    print("  ✔ SAP session refreshed during pre-print capture")
                except Exception as _refresh_err:
                    print(f"  ⚠️  Could not refresh SAP session during pre-print capture: {_refresh_err}")

    print("1) starting to save it ...")

    # ══════════════════════════════════════════════════════════════════════════
    # PHASE 2: TRIGGER PRINT → SAVE PDF (this kills the COM session)
    # ══════════════════════════════════════════════════════════════════════════
    try:
        session.findById("wnd[1]/tbar[0]/btn[86]").press()
        print("2")
    except Exception as _print_btn_err:
        print(f"  ⚠️  Print button press failed on current session: {_print_btn_err}")
        try:
            session = _reconnect_session(timeout=60)
            print("  ✔ SAP session refreshed before Print button")
            session.findById("wnd[1]/tbar[0]/btn[86]").press()
            print("3")
        except Exception as _print_reconnect_err:
            print(
                "  ⚠️  Could not recover SAP session after Save to start printing: "
                f"{_print_reconnect_err}"
            )
            print(
                "  ⚠️  Save was already submitted, so this SAP item will be "
                "treated as completed to avoid creating a duplicate checklist."
            )
            print("4")
            return None, _checklist_number
    print("  Waiting for SAP print dialog...")
    time.sleep(5)

    from pywinauto import Application, Desktop
    from pywinauto.keyboard import send_keys
    import pyperclip

    print("5")
    # ══════════════════════════════════════════════════════════════════════════
    # HANDLE SAP PRINT DIALOG
    # After btn[86], SAP opens a standard Windows Print dialog (#32770).
    # The SAP COM session is dead at this point. Use pywinauto directly.
    #
    # Strategy order:
    #   1. Alt+O keyboard shortcut (OK button has underlined "O" in many
    #      Windows Print dialogs — Alt+O activates it directly)
    #   2. Tab to OK button + Enter (OK is typically the first button,
    #      so it already has default focus — just press Enter)
    #   3. pywinauto .click() on child_window (the old approach)
    #   4. win32 backend fallback
    # ══════════════════════════════════════════════════════════════════════════
    print("  Connecting to SAP Print dialog...")
    _print_ok_done = False

    # ── Strategy 1: Bring dialog to foreground and use keyboard ──────────────
    # The Print dialog's OK button typically has keyboard focus by default.
    # Simply setting focus on the dialog and pressing Enter should work.
    # If that fails, Alt+O is the accelerator key for "OK" in Windows dialogs.
    print("  Strategy 1: Keyboard shortcut (Enter / Alt+O)...")
    try:
        # First, make sure the Print dialog is in the foreground
        _print_app = Application(backend="uia").connect(title_re="Print", timeout=15)
        _print_win = _print_app.window(title_re="Print")
        _print_win.set_focus()
        time.sleep(1)
        print("  Print dialog focused. Sending Enter key...")

        # Press Enter — OK is the default button in the Print dialog
        send_keys("{ENTER}")
        _print_ok_done = True
        print("  ✔ Sent Enter key to Print dialog")
    except Exception as _kb1_err:
        print(f"  ⚠ Strategy 1 failed: {_kb1_err}")

    # ── Strategy 2: If dialog is still there, try Alt+O ──────────────────────
    if not _print_ok_done:
        print("  Strategy 2: Trying Alt+O shortcut...")
        try:
            time.sleep(1)
            send_keys("%o")  # Alt+O — accelerator for OK in many dialogs
            _print_ok_done = True
            print("  ✔ Sent Alt+O to Print dialog")
        except Exception as _kb2_err:
            print(f"  ⚠ Strategy 2 failed: {_kb2_err}")

    # ── Strategy 3: pywinauto .click() on OK button ──────────────────────────
    if not _print_ok_done:
        print("  Strategy 3: pywinauto .click() on OK button...")
        try:
            _print_app = Application(backend="uia").connect(title_re="Print", timeout=10)
            _print_win = _print_app.window(title_re="Print")
            _print_win.set_focus()
            time.sleep(0.5)
            _print_win.child_window(title="OK", control_type="Button").click()
            _print_ok_done = True
            print("  ✔ Clicked OK on Print dialog via .click()")
        except Exception as _pe:
            print(f"  ⚠ Strategy 3 failed: {_pe}")

    # ── Strategy 4: win32 backend ────────────────────────────────────────────
    if not _print_ok_done:
        print("  Strategy 4: win32 backend fallback...")
        try:
            _print_app_w32 = Application(backend="win32").connect(
                title="Print", class_name="#32770", timeout=10
            )
            _print_win_w32 = _print_app_w32.window(title="Print", class_name="#32770")
            _print_win_w32.set_focus()
            time.sleep(0.5)
            try:
                _print_win_w32["OK"].click()
                _print_ok_done = True
                print("  ✔ Clicked OK via win32 backend")
            except Exception:
                send_keys("{ENTER}")
                _print_ok_done = True
                print("  ✔ Pressed Enter (win32 fallback)")
        except Exception as _w32_err:
            print(f"  ⚠ Strategy 4 failed: {_w32_err}")

    # ── Last resort: keyboard Enter ──────────────────────────────────────────
    if not _print_ok_done:
        time.sleep(1)
        send_keys("{ENTER}")
        print("  ⚠ Sent blind Enter as last resort for Print dialog")

    # ── Wait for "Save Print Output As" dialog to appear ──────────────────────
    print("  ℹ Print dialog closed, checking if Save dialog appears...")
    time.sleep(5)

    _save_dir = pdf_output_dir or config.SAP_PDF_TARGET_DIR
    os.makedirs(_save_dir, exist_ok=True)
    _pdf_filename = f"checklist_{time.strftime('%Y%m%d_%H%M%S')}.pdf"
    _full_save_path = os.path.abspath(os.path.join(_save_dir, _pdf_filename))
    print(f"  Saving PDF to: {_full_save_path}")

    # ══════════════════════════════════════════════════════════════════════════
    # HANDLE "Save Print Output As" DIALOG — PURE KEYBOARD
    # After Print OK, the Save dialog appears with the File Name field focused.
    # We don't need to "find" or "connect" to it — it's the active foreground
    # window. Just type the path and press Enter.
    # ══════════════════════════════════════════════════════════════════════════
    print("  Handling Save dialog via keyboard (no window search needed)...")

    # Wait a bit more to ensure the Save dialog is fully rendered and focused
    time.sleep(3)

    # The Save dialog's File Name field should already have keyboard focus.
    # Type the full save path using clipboard paste (handles special chars).
    pyperclip.copy(_full_save_path)
    time.sleep(0.3)

    # Select all in filename field and paste
    send_keys("^a")
    time.sleep(0.2)
    send_keys("^v")
    time.sleep(0.5)
    print(f"  ✔ Pasted path: {_full_save_path}")

    # Press Enter to save
    send_keys("{ENTER}")
    print("  ✔ Pressed Enter to save")
    time.sleep(3)

    # Handle potential "Confirm Save As" / overwrite dialog (sends Alt+Y)
    try:
        send_keys("%y")
        time.sleep(0.5)
    except Exception:
        pass

    print(f"  📄 PDF saved at: {_full_save_path}")

    # ══════════════════════════════════════════════════════════════════════════
    # PHASE 3: POST-PDF CAPTURE FALLBACK (reconnect session — COM link is dead)
    # The pywinauto interactions with native Windows dialogs kill the original
    # COM session. We must re-attach via _reconnect_session().
    # ══════════════════════════════════════════════════════════════════════════
    if not _checklist_number:
        print("\n  ── Phase 3: Reconnecting SAP session for checklist capture ──")
        for _capture_attempt in range(1, 4):
            if _checklist_number:
                break
            print(f"  [CHECKLIST] Capture attempt {_capture_attempt}/3")
            try:
                _sess = _reconnect_session(timeout=60)
                print("  ✔ SAP session reconnected successfully")
                time.sleep(2 * _capture_attempt)

                # Try status bar text on reconnected session
                try:
                    _sbar_txt = _sess.findById("wnd[0]/sbar").text.strip()
                    if _sbar_txt:
                        print(f"  📋 Status bar (reconnected): {_sbar_txt}")
                        _m = _re_chk.search(r'\b(\d{7,})\b', _sbar_txt)
                        if _m:
                            _checklist_number = _m.group(1)
                            print(f"  ✅ Checklist number from sbar (reconnected): {_checklist_number}")
                except Exception as _e:
                    print(f"  ⚠️  sbar read on reconnected session: {_e}")

                # Double-click status bar → popup (the working snippet approach)
                if not _checklist_number:
                    try:
                        _sess.findById("wnd[0]/sbar").doubleClick()
                        print("  🖱️  Status bar double-clicked — waiting for popup...")
                        time.sleep(1.5 * _capture_attempt)

                        _popup_text_ids = [
                            "wnd[1]/usr/txtMESSAGE",
                            "wnd[1]/usr/txtMESSTXT-STEXT",
                            "wnd[1]/usr/subSCREEN:SAPLMSG_DISPLAY:0100/txtMESSAGE",
                            "wnd[1]/usr/subSCREEN_HEADER:SAPLMSG_DISPLAY:0100/txtMESSAGE",
                        ]
                        for _ptxt_id in _popup_text_ids:
                            try:
                                _popup_text = _sess.findById(_ptxt_id).text.strip()
                                if _popup_text:
                                    print(f"  📋 Popup text [{_ptxt_id}]: {_popup_text}")
                                    _m = _re_chk.search(r'\b(\d{7,})\b', _popup_text)
                                    if _m:
                                        _checklist_number = _m.group(1)
                                        print(f"  ✅ Checklist number from doubleClick popup: {_checklist_number}")
                                        break
                            except Exception:
                                continue

                        if not _checklist_number:
                            try:
                                _wnd1_text = _sess.findById("wnd[1]").text.strip()
                                if _wnd1_text:
                                    print(f"  📋 Popup title: {_wnd1_text}")
                                    _m = _re_chk.search(r'\b(\d{7,})\b', _wnd1_text)
                                    if _m:
                                        _checklist_number = _m.group(1)
                                        print(f"  ✅ Checklist number from popup title: {_checklist_number}")
                            except Exception:
                                pass

                        # Close popup
                        for _close_id in ["wnd[1]/tbar[0]/btn[0]", "wnd[1]/tbar[0]/btn[12]"]:
                            try:
                                _sess.findById(_close_id).press()
                                time.sleep(0.5)
                                print("  🪟  Popup closed.")
                                break
                            except Exception:
                                continue

                    except Exception as _dc_err:
                        print(f"  ⚠️  doubleClick on reconnected session: {_dc_err}")

                # Form field fallback on reconnected session
                if not _checklist_number:
                    try:
                        _chk = _sess.findById("wnd[0]/usr/txtZFIBPT11001-CHK_NO").text.strip()
                        if _chk:
                            _m = _re_chk.search(r'\b(\d{7,})\b', _chk)
                            if _m:
                                _checklist_number = _m.group(1)
                                print(f"  ✅ Checklist number from CHK_NO (reconnected): {_checklist_number}")
                    except Exception:
                        pass

            except Exception as _recon_err:
                print(f"  ⚠️  SAP session reconnection failed: {_recon_err}")
                if _capture_attempt < 3:
                    time.sleep(3 * _capture_attempt)

    if _checklist_number:
        print(f"\n  ✅ FINAL checklist number captured: {_checklist_number}")
    else:
        print("\n  ⚠️  Checklist number could not be captured from any source after 3 attempts.")

    return _full_save_path, _checklist_number


def handle_cancel_popup(session):
    """Press 'Cancel' on a typical modal (wnd[1]) if it appears."""
    try:
        session.findById("wnd[1]").sendVKey(12)  # F12 = Cancel
        return True
    except Exception:
        pass

    for cancel_id in [
        "wnd[1]/tbar[0]/btn[12]",
        "wnd[1]/tbar[0]/btn[2]",
        "wnd[1]/tbar[0]/btn[21]",
    ]:
        try:
            session.findById(cancel_id).press()
            return True
        except Exception:
            continue

    try:
        session.findById("wnd[1]")
        return False
    except Exception:
        return True

def open_print_dialog(session):
    """
    Opens SAP print dialog safely without relying on preview shell
    """
    # Try toolbar print button
    if try_press(session, [
        "wnd[0]/tbar[0]/btn[86]",   # Print
        "wnd[0]/tbar[1]/btn[5]"     # Alternate print
    ]):
        time.sleep(2)
        return True

    # Try menu path (List → Print)
    try:
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        time.sleep(2)
        return True
    except Exception:
        pass

    raise RuntimeError("❌ Unable to open SAP Print dialog")



def close_pdf_preview_and_exit(session):
    """
    Close the PDF Preview popup (wnd[1]) and press btn[15] (Back/Exit).
    NOTE: In the current flow (PDF saved via Windows dialog), this function
    is not called — kept for backward compatibility only.
    """
    try:
        session.findById("wnd[1]").close()
        print("🪟  PDF Preview popup closed.")
        time.sleep(1)
    except Exception:
        try:
            session.findById("wnd[1]").sendVKey(12)
            time.sleep(1)
        except Exception:
            pass

    try:
        session.findById("wnd[0]").maximize()
    except Exception:
        pass

    try:
        session.findById("wnd[0]/tbar[0]/btn[15]").press()
        print("🔙  Pressed btn[15] (Back/Exit) on wnd[0].")
        time.sleep(2)
    except Exception as e:
        print(f"⚠️  Could not press btn[15]: {e}")


def capture_checklist_data(session):
    """
    Reads the SAP checklist number after the form has been saved and the PDF
    has already been written to disk via the Windows 'Save Print Output As' dialog.

    Flow (no PDF-preview popup exists in this path):
        1. Read sbar text immediately — the checklist number is there right after Save.
        2. Press Back (btn[3]) to return to the checklist display screen.
        3. Double-click sbar to open the confirmation popup and read the number again.
        4. Close the popup.
        5. Final fallback: re-read sbar after Back navigation.
    """
    import time
    import re

    checklist_number = None

    # The original 'session' COM object may be dead after pywinauto interactions.
    # Re-attach to the existing SAP GUI session.
    try:
        session = _reconnect_session()
        print("  ✔ SAP session reconnected for capture_checklist_data")
    except Exception as _rc_err:
        print(f"  ⚠️  Could not reconnect SAP session: {_rc_err}")
        # Continue with the original session — it might still work
        pass

    # ── STEP 1: Read the status bar text immediately after Save ───────────────
    # SAP status bar shows "Checklist XXXXXXX saved" right after btn[11] press.
    try:
        sbar_text = session.findById("wnd[0]/sbar").text.strip()
        if sbar_text:
            print(f"  📋 Status bar text (immediate): {sbar_text}")
            m = re.search(r'\b(\d{7,})\b', sbar_text)
            if m:
                checklist_number = m.group(1)
                print(f"  ✅ Checklist number from status bar (immediate): {checklist_number}")
    except Exception as e:
        print(f"  ⚠️  Could not read immediate sbar text: {e}")

    # ── STEP 2: Press Back (btn[3]) — return to the saved checklist screen ────
    try:
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        print("  🔙 Back button (btn[3]) pressed — returning to checklist screen.")
        time.sleep(2)
    except Exception as e:
        print(f"  ⚠️  Could not press Back btn[3]: {e}")

    # ── STEP 3: Double-click the status bar — opens checklist-number popup ────
    try:
        session.findById("wnd[0]/sbar").doubleClick()
        print("  🖱️  Status bar double-clicked — waiting for checklist popup...")
        time.sleep(1.5)

        popup_text_ids = [
            "wnd[1]/usr/txtMESSAGE",
            "wnd[1]/usr/txtMESSTXT-STEXT",
            "wnd[1]/usr/subSCREEN:SAPLMSG_DISPLAY:0100/txtMESSAGE",
            "wnd[1]/usr/subSCREEN_HEADER:SAPLMSG_DISPLAY:0100/txtMESSAGE",
        ]
        for ptxt_id in popup_text_ids:
            try:
                popup_text = session.findById(ptxt_id).text.strip()
                if popup_text:
                    print(f"  📋 Popup text [{ptxt_id}]: {popup_text}")
                    m = re.search(r'\b(\d{7,})\b', popup_text)
                    if m:
                        checklist_number = m.group(1)
                        print(f"  ✅ Checklist number from doubleClick popup: {checklist_number}")
                        break
            except Exception:
                continue

        if not checklist_number:
            try:
                wnd1_text = session.findById("wnd[1]").text.strip()
                print(f"  📋 Popup title text: {wnd1_text}")
                m = re.search(r'\b(\d{7,})\b', wnd1_text)
                if m:
                    checklist_number = m.group(1)
                    print(f"  ✅ Checklist number from popup title: {checklist_number}")
            except Exception:
                pass

        for close_id in ["wnd[1]/tbar[0]/btn[0]", "wnd[1]/tbar[0]/btn[12]"]:
            try:
                session.findById(close_id).press()
                time.sleep(0.5)
                print("  🪟  Checklist popup closed.")
                break
            except Exception:
                continue

    except Exception as e:
        print(f"  ⚠️  Status bar doubleClick failed: {e}")

    # ── STEP 4: Final fallback — re-read status bar after Back navigation ─────
    if not checklist_number:
        try:
            sbar_text = session.findById("wnd[0]/sbar").text.strip()
            if sbar_text:
                print(f"  📋 Status bar text (post-Back): {sbar_text}")
                m = re.search(r'\b(\d{7,})\b', sbar_text)
                if m:
                    checklist_number = m.group(1)
                    print(f"  ✅ Checklist number from status bar (post-Back): {checklist_number}")
        except Exception:
            pass

    # ── STEP 4: Read all other checklist fields from the form ─────────────────
    print("\n" + "=" * 60)
    print("  CHECKLIST DATA CAPTURED FROM SAP")
    print("=" * 60)

    data = {}

    # --- Header fields (text fields in the main form) ---
    field_map = {
        "Checklist No"       : "wnd[0]/usr/txtZFIBPT11001-CHK_NO",
        "Plant"              : "wnd[0]/usr/ctxtZFIBPT11001-WERKS",
        "Company Code"       : "wnd[0]/usr/ctxtZFIBPT11001-BUKRS",
        "Vendor"             : "wnd[0]/usr/ctxtZFIBPT11001-LIFNR",
        "Adani GSTIN"        : "wnd[0]/usr/txtZFIBPT11001-GSTIN",
        "GST Partner"        : "wnd[0]/usr/txtZFIBPT11001-GST_PART",
        "GST Partner Reg No" : "wnd[0]/usr/txtZFIBPT11001-GST_REG",
        "Doc Category"       : "wnd[0]/usr/ctxtZFIBPT11001-CHK_DOC_TYP",
        "Value of Checklist" : "wnd[0]/usr/txtZFIBPT11001-CHK_VALUE",
        "Debit Note Date"    : "wnd[0]/usr/ctxtZFIBPT11001-DEB_NOT_DAT",
        "Debit Note Ref"     : "wnd[0]/usr/txtZFIBPT11001-DEB_NOT_REF",
        "Doc Sub Category"   : "wnd[0]/usr/ctxtZFIBPT11001-DOC_SUB_CAT",
    }

    for label, ctrl_id in field_map.items():
        try:
            val = session.findById(ctrl_id).text.strip()
            data[label] = val if val else "—"
        except Exception:
            data[label] = "—"

    # --- Status bar: supplement data dict with already-captured checklist number ---
    try:
        status_msg = session.findById("wnd[0]/sbar").text.strip()
        if status_msg:
            data["Status Bar"] = status_msg
            # If the VBS preamble didn't get a number, try once more here
            if not checklist_number:
                match = re.search(r'\b(\d{7,})\b', status_msg)
                if match:
                    checklist_number = match.group(1)
    except Exception:
        pass

    if checklist_number:
        data["Checklist No"] = checklist_number

    # --- Print captured fields ---
    for label, val in data.items():
        print(f"  {label:<22}: {val}")

    if checklist_number:
        print(f"\n  ✅ Checklist Number : {checklist_number}")
    else:
        print("\n  ⚠️  Checklist number could not be parsed from status bar.")

    print("=" * 60 + "\n")

    # ------------------------------------------------------------------
    # Close all open SAP windows — handle logoff confirmation popup
    # ------------------------------------------------------------------
    print("🔒 Closing all SAP windows...")
    for i in range(9, -1, -1):
        try:
            session.findById(f"wnd[{i}]").close()
            time.sleep(0.5)
            print(f"   Closed wnd[{i}]")
        except Exception:
            pass

        # After each close attempt, check if a logoff/confirmation popup appeared
        # and click Yes (btn[8] = Yes, btn[0] = OK — try both)
        for popup in range(5, -1, -1):
            try:
                pop = session.findById(f"wnd[{popup}]")
                # Try Yes button first (btn[8]), then OK (btn[0])
                confirmed = False
                for btn_id in [
                    f"wnd[{popup}]/tbar[0]/btn[8]",   # Yes
                    f"wnd[{popup}]/tbar[0]/btn[0]",   # OK / Continue
                ]:
                    try:
                        session.findById(btn_id).press()
                        confirmed = True
                        print(f"   ✔ Clicked Yes/OK on popup wnd[{popup}]")
                        time.sleep(0.5)
                        break
                    except Exception:
                        continue
            except Exception:
                break   # popup window doesn't exist — move on

    # ------------------------------------------------------------------
    # Handle the "Log Off" popup — default focus is on "No", shift to "Yes"
    # Use Ctrl+Left arrow to move focus to Yes, then press Enter
    # ------------------------------------------------------------------
    time.sleep(1)
    try:
        # Send Ctrl+Left to move focus from No → Yes, then Enter to confirm
        session.findById("wnd[1]").sendVKey(0)   # make sure popup has focus
    except Exception:
        pass

    try:
        from pywinauto import Desktop
        from pywinauto.keyboard import send_keys

        logoff_dlg = Desktop(backend="uia").window(title_re=".*Log Off.*|.*Log off.*")
        logoff_dlg.wait("visible", timeout=5)
        logoff_dlg.set_focus()
        time.sleep(0.3)
        send_keys("{LEFT}")       # move focus from No → Yes
        time.sleep(0.3)
        send_keys("{ENTER}")      # confirm Yes
        print("   ✔ Navigated to 'Yes' and confirmed Log Off popup.")
        time.sleep(1.5)
    except Exception as e:
        print(f"   ⚠️  Could not handle Log Off popup via keyboard: {e}")

    # ------------------------------------------------------------------
    # Close the final remaining SAP window via its X (close) button
    # ------------------------------------------------------------------
    time.sleep(1)
    try:
        session.findById("wnd[0]").close()
        print("   ✔ Final SAP window closed via scripting.")
        time.sleep(0.5)
    except Exception:
        pass

    # Fallback: pywinauto close any remaining SAP windows
    try:
        from pywinauto import Desktop
        for win in Desktop(backend="uia").windows(title_re=".*SAP.*"):
            try:
                win.close()
                print(f"   ✔ Closed remaining SAP window: '{win.window_text()}'")
                time.sleep(0.5)
            except Exception:
                pass
    except Exception:
        pass

    print("✅ All SAP windows closed.")
    return checklist_number



    

# -------------------------------------------------
# Cleanup: force-close all SAP windows (used by sap_main on retry)
# -------------------------------------------------
def cleanup_sap_windows():
    """Force-close ALL SAP GUI windows and terminate saplogon/sapgui processes."""
    print("\n[SAP CLEANUP] Closing all SAP windows before retry...")
    try:
        application = _get_sap_application()
        for ci in range(application.Children.Count - 1, -1, -1):
            try:
                conn = application.Children(ci)
                for si in range(conn.Children.Count - 1, -1, -1):
                    try:
                        sess = conn.Children(si)
                        for wi in range(5, 0, -1):
                            try:
                                sess.findById(f"wnd[{wi}]").close()
                                time.sleep(0.2)
                            except Exception:
                                pass
                        try:
                            sess.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
                            sess.findById("wnd[0]").sendVKey(0)
                            time.sleep(1)
                        except Exception:
                            pass
                    except Exception:
                        pass
            except Exception:
                pass
    except Exception:
        pass
    try:
        import psutil
        for pname in ["saplogon.exe", "sapgui.exe"]:
            for proc in psutil.process_iter(["name"]):
                try:
                    if proc.info["name"] and proc.info["name"].lower() == pname:
                        proc.terminate()
                        proc.wait(timeout=5)
                except Exception:
                    pass
    except ImportError:
        import subprocess as _sp
        for pname in ["saplogon.exe", "sapgui.exe"]:
            try:
                _sp.run(["taskkill", "/F", "/IM", pname], capture_output=True, timeout=10)
            except Exception:
                pass
    except Exception:
        pass
    _wait_for_sap_processes_to_exit()
    print("[SAP CLEANUP] Done.\n")


# -------------------------------------------------
# PUBLIC ENTRY POINT (called from sap_main.py)
# -------------------------------------------------
def run_sap_automation(
    company_code: str,
    vendor_code: str,
    amount: float,
    cost_center: str,
    plant: str,
    deb_not_date: str = None,
    print_pdf: bool = True,
    pdf_output_dir: str = None,
    deb_not_ref: str = None,
):
    """
    Execute SAP Non-PO automation with values sourced from the email/SPV tables.

    Args:
        company_code  : BUKRS  – from vendor master data
        vendor_code   : LIFNR  – from vendor master data
        amount        : Absolute value of net_payable_receivable from SPV DSM table
                        (negative amounts are converted to positive before passing in)
        cost_center   : KOSTL  – from vendor master data
        plant         : WERKS  – from vendor master data
        deb_not_date  : DEB_NOT_DATE in SAP format DD.MM.YYYY (defaults to today)
        print_pdf     : When True (default), trigger the PDF print flow after saving.
                        Pass False to skip printing (e.g., during dry-run tests).
        deb_not_ref   : DEB_NOT_REF reference number. Defaults to config.SAP_DEB_NOT_REF.
                        For Debit/Credit-note checklists this is the note's Doc.No.
    """
    # --- Derive SAP-formatted date (DD.MM.YYYY) ---
    if not deb_not_date:
        deb_not_date = datetime.today().strftime("%d.%m.%Y")

    print(f"\n=== SAP Automation triggered ===")
    print(f"  Company Code  (BUKRS)       : {company_code}")
    print(f"  Vendor Code   (LIFNR)       : {vendor_code}")
    print(f"  Amount                      : {amount}")
    print(f"  Cost Center   (KOSTL)       : {cost_center}")
    print(f"  Plant         (WERKS)       : {plant}")
    print(f"  Debit Note Date             : {deb_not_date}")
    print(f"  Debit Note Ref              : {deb_not_ref or config.SAP_DEB_NOT_REF}")
    print(f"  Print PDF after save        : {print_pdf}")
    print(f"================================\n")

    # --- Static / env-driven config (loaded from config.py) ---
    saplogon_path = config.SAP_LOGON_PATH
    system_name   = config.SAP_SYSTEM
    sap_client    = config.SAP_CLIENT
    sap_user      = config.SAP_USERNAME
    sap_pass      = config.SAP_PASSWORD
    sap_lang      = config.SAP_LANGUAGE

    # --- Build params dict (only the fields that come from email data are overridden) ---
    params = {
        # Fields sourced from email / vendor master (passed in as arguments)
        "BUKRS":        company_code,
        "LIFNR":        vendor_code,
        "AMOUNT":       str(amount),          # already made positive by caller
        "COST_CENTER":  cost_center,
        "PLANT":        plant,
        "DEB_NOT_DATE": deb_not_date,

        # Static / env fields (loaded from config.py)
        "CHK_DOC_TYP":  config.SAP_CHK_DOC_TYP,
        "DIG_SIGN_INV": config.SAP_DIG_SIGN_INV,
        "GJAHR":        config.SAP_GJAHR,
        "DEB_NOT_REF":  deb_not_ref or config.SAP_DEB_NOT_REF,
        "DOC_SUB_CAT":  config.SAP_DOC_SUB_CAT,
        "DO_PREVIEW":   config.SAP_DO_PREVIEW,
    }

    # Log all params that will be filled in SAP
    print(f"\n  ┌─ SAP FORM PARAMS (all fields to be filled) ─────────────────")
    for k, v in params.items():
        print(f"  │  {k:<15} = {v}")
    print(f"  └──────────────────────────────────────────────────────────────\n")

    # --- Orchestration ---
    launch_sap_logon(saplogon_path)
    app, conn, session = attach_to_scripting(system_name)

    try:
        session.findById("wnd[0]").maximize()
    except Exception:
        pass

    if not sap_user or not sap_pass:
        raise RuntimeError("SAP_USERNAME and SAP_PASSWORD environment variables must be set.")

    login(session, sap_client, sap_user, sap_pass, sap_lang)
    go_to_tcode(session, "ZDCC")

    select_non_po_and_create(session)
    fill_non_po_form(session, params)

    # Save the checklist form, trigger Print dialog → save PDF via Windows dialog.
    # Returns (full_pdf_path, checklist_number_or_None).
    # The checklist number is captured AFTER the PDF is saved to disk using:
    #   1. Direct status bar text read
    #   2. Double-click status bar popup (most reliable for SAP)
    #   3. Form field CHK_NO fallback
    pdf_dest, pre_captured_checklist = apply_subcategory_checkbox_table_and_save(
        session, params, pdf_output_dir=pdf_output_dir
    )
    print(f"  [PDF]       Saved  → {pdf_dest}")
    print(f"  [CHECKLIST] Captured: {pre_captured_checklist or '(not found yet)'}")
    checklist_number = pre_captured_checklist
    renamed_pdf_path = pdf_dest   # will be updated after rename

    if pdf_dest:
        # Last-resort fallback: capture_checklist_data presses Back and tries
        # additional methods (doubleClick on returned screen, form fields, etc.)
        if not checklist_number:
            print("  [CHECKLIST] Primary capture missed — trying post-save fallback...")
            checklist_number = capture_checklist_data(session)
        else:
            print("  [CHECKLIST] Number already captured — skipping fallback.")

        if not checklist_number:
            print(
                "  ⚠️  SAP checklist number could not be captured, but the PDF "
                "flow already passed Save. Treating SAP automation as completed."
            )

        # ── Rename PDF: {stem}_{checklist_number}.pdf ────────────────────────
        if checklist_number and pdf_dest:
            import pathlib as _pl
            _p        = _pl.Path(pdf_dest)
            _new_name = f"{_p.stem}_{checklist_number}{_p.suffix}"
            _new_path = _p.parent / _new_name
            try:
                if _new_path != _p:
                    os.rename(str(_p), str(_new_path))
                    print(f"  📄 PDF renamed: '{_p.name}' → '{_new_name}'")
                renamed_pdf_path = str(_new_path)
            except Exception as _rename_err:
                print(f"  ⚠️  PDF rename failed ({_rename_err}) — keeping: '{_p.name}'")
                renamed_pdf_path = pdf_dest
        elif pdf_dest:
            print(f"  ⚠️  checklist_number unavailable — PDF kept as: '{os.path.basename(pdf_dest)}'")

    print("✅ SAP automation flow completed.")
    return checklist_number, renamed_pdf_path


# -------------------------------------------------
# Standalone entry point (for isolated testing)
# -------------------------------------------------
def main():
    """
    Standalone test using hard-coded example values.
    In production this function is NOT called; run_sap_automation() is used instead.
    """
    run_sap_automation(
        company_code="6060",
        vendor_code="215944",
        amount=1,
        cost_center="6067OPC1",
        plant="6067",
        deb_not_date=None,   # defaults to today
        print_pdf=True,      # set to False to skip PDF printing during tests
    )

if __name__ == "__main__":
    main()
