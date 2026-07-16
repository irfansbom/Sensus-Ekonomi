from pathlib import Path
from playwright.sync_api import sync_playwright
from playwright_stealth import Stealth
Path("download").mkdir(exist_ok=True)
with sync_playwright() as p:
    browser = p.chromium.launch(
        headless=False, args=["--disable-blink-features=AutomationControlled"]
    )
    # context = browser.new_context(
    #     storage_state="Dashboard Scrapper/session_dash.json", accept_downloads=True
    # )
    context = browser.new_context(
        storage_state="Dashboard Scrapper/session_dash.json",
        accept_downloads=True,
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
        viewport={"width": 1920, "height": 1000},
        locale="id-ID",
        timezone_id="Asia/Jakarta",
    )
    page = context.new_page()
    stealth = Stealth()
    stealth.apply_stealth_sync(page)  # inject sebelum goto

    page.goto("https://dashboard-se2026.apps.bps.go.id")

    while True:
        key = input(
            "\n[SPASI] Download Kabupaten/Kota"
            "\n[d]     Download Mikro"
            "\n[h]     Download Laporan Harian"
            "\n[Enter] Keluar"
            "\nPilihan: "
        )
        # Keluar
        if key == "":
            break
        try:
            if key == " ":
                with page.expect_download(timeout=600000) as d:
                    page.get_by_role(
                        "menuitem",
                        name="Sampai Kabupaten / Kota"
                    ).click()

                download = d.value
            elif key.lower() == "d":
                with page.expect_download(timeout=600000) as d:
                    page.get_by_role(
                        "button",
                        name="Unduh Mikro"
                    ).click()

                download = d.value
            elif key.lower() == "h":
                with page.expect_download(timeout=600000) as d:
                    page.get_by_role("button", name="Download XLSX").click()

                download = d.value
            else:
                print("Input tidak dikenali.")
                continue

            file_path = Path("download") / download.suggested_filename
            download.save_as(file_path)
            print(f"✓ Download selesai: {file_path}")
        except Exception as e:
            print("Gagal download:", e)
    browser.close()
