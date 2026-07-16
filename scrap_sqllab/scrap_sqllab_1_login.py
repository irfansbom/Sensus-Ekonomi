from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://fasih-dashboard.bps.go.id/superset/sqllab/")
    input("Login manual lalu tekan Enter...")
    context.storage_state(path="./session_sqllab.json")
    # browser.close()
