# from playwright.sync_api import sync_playwright

# with sync_playwright() as p:
#     browser = p.chromium.launch(headless=False)
#     context = browser.new_context()
#     page = context.new_page()
#     page.goto("https://dashboard-se2026.apps.bps.go.id/login")
#     input("Login manual lalu tekan Enter...")
#     context.storage_state(path="./Dashboard Scrapper/session_dash.json")
#     # browser.close()


from playwright.sync_api import sync_playwright
from playwright_stealth import Stealth
import json

with sync_playwright() as p:
    browser = p.chromium.launch(
        headless=False, args=["--disable-blink-features=AutomationControlled"]
    )
    context = browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
        viewport={"width": 1366, "height": 768},
        locale="id-ID",
        timezone_id="Asia/Jakarta",
    )
    page = context.new_page()

    stealth = Stealth()
    stealth.apply_stealth_sync(page)  # inject sebelum goto

    page.goto("https://dashboard-se2026.apps.bps.go.id/login")

    input("Login manual lalu tekan Enter...")
    context.storage_state(path="./Dashboard Scrapper/session_dash.json")
    cookies = context.cookies()
    with open("./Dashboard Scrapper/cookie_dash.json", "w") as f:
        json.dump(cookies, f, indent=2)
# browser.close()


# from playwright.sync_api import sync_playwright
# import json

# with sync_playwright() as p:
#     browser = p.chromium.launch(headless=False)
#     context = browser.new_context()
#     page = context.new_page()
#     page.goto("https://dashboard-se2026.apps.bps.go.id/login")
#     input("Login manual lalu tekan Enter...")
#     context.storage_state(path="./Dashboard Scrapper/session_dash.json")
#     cookies = context.cookies()
#     with open("./Dashboard Scrapper/cookie_dash.json", "w") as f:
#         json.dump(cookies, f, indent=2)

#     browser.close()
