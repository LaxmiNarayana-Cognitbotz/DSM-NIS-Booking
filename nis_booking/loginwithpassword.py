from playwright.sync_api import sync_playwright

def run(playwright):
    edge_path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

    browser = playwright.chromium.launch(
        headless=False,
        executable_path=edge_path
    )
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://ael-nis-prd.launchpad.cfapps.in30.hana.ondemand.com/site?siteId=6b54bd07-1384-49b1-a4e9-781f4f23eef1#Shell-home")

    # Step 1: Fill email
    page.wait_for_selector("#i0116")
    page.fill("#i0116", "agelnis001@adani.com")
    page.click("#idSIButton9")

    # Step 2: Fill password
    page.wait_for_selector("#i0118")
    page.fill("#i0118", "AGEL.NIS@two02six")
    page.click("#idSIButton9")

    # Optional: wait for navigation after login
    page.wait_for_load_state("networkidle")

    # Step 3: Click on the ExternalAuth card (Approve with RSAEntraMFA)
    page.wait_for_selector("div[data-value='ExternalAuth']")
    page.click("div[data-value='ExternalAuth']")


    
    # Step 4: Fill RSA OTP input and submit
    try:
        # Wait for the OTP input to be visible and enabled
        page.wait_for_selector('#input_otp_secret', state='visible', timeout=20000)
        # Fill with 1001 (as requested)
        page.fill('#input_otp_secret', '1001')
        # Ensure the Submit button is enabled before clicking
        page.wait_for_selector('#btn_verify_securid:not([aria-disabled="true"])', state='attached', timeout=10000)
        page.click('#btn_verify_securid')
        # Optional: wait for navigation or network to settle
        page.wait_for_load_state('networkidle')
    except Exception as e:
        print(f'OTP step failed: {e}')

    # Step 5: Click the final "Yes" / continue submit button
    try:
        # Wait until the button is visible & enabled (isPrimaryButtonEnabled)
        page.wait_for_selector('#idSIButton9', state='visible', timeout=20000)
        # Some tenants toggle enablement via aria-describedby/KMsI/knockout binding; add a small wait for enabling
        with page.expect_navigation(wait_until='networkidle'):
            page.click('#idSIButton9')
    except Exception as e:
        print(f'Final submit step failed: {e}')
    print("Login flow completed, ExternalAuth card and Yes button clicked")
    input("Press Enter to close...")
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
