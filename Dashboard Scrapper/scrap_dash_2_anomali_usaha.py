from playwright.sync_api import sync_playwright
import pandas as pd
import os
import time
from datetime import datetime
from playwright_stealth import Stealth

kode_kab_list = [
    # "1601",
    # "1602",
    # "1603",
    # "1604",
    # "1605",
    # "1606",
    # "1607",
    # "1608",
    # "1609",
    "1610",
    # "1611",
    # "1612",
    # "1613",
    # "1671",
    # "1672",
    # "1673",
    # "1674",
]

all_df = []

with sync_playwright() as p:

    browser = p.chromium.launch(
        headless=False, args=["--disable-blink-features=AutomationControlled"]
    )

    context = browser.new_context(
        storage_state="Dashboard Scrapper/session_dash.json",
        accept_downloads=True,
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
        viewport={"width": 1366, "height": 768},
        locale="id-ID",
        timezone_id="Asia/Jakarta",
    )
    page = context.new_page()
    stealth = Stealth()
    stealth.apply_stealth_sync(page)  # inject sebelum goto

    page.goto("https://dashboard-se2026.apps.bps.go.id")
    input("Selesaikan captcha dulu, lalu tekan ENTER...")
    for kode_kab in kode_kab_list:

        print(f"Proses {kode_kab}")

        try:
            response = context.request.get(
                "https://dashboard-se2026.apps.bps.go.id/api/mikro/anomali-case-kab",
                params={
                    "kode_kabupaten": kode_kab,
                    "indikator": "128,129,130,131,132,133,134,135",
                    "sudah_indikator": "40,41,42,43,44,45,46,46",
                    "type": "usaha",
                    "anomali_no": "1,2,3,4,5,6,7,8",
                },
            )

            print("Status:", response.status)
            print("Content-Type:", response.headers.get("content-type"))

            if "application/json" not in response.headers.get("content-type", ""):
                print(response.text()[:1000])
                continue

            data = response.json()

            # sesuaikan dengan struktur JSON
            df = pd.DataFrame(data)

            # tandai asal kabupaten
            df["kode_kabupaten"] = kode_kab

            all_df.append(df)

            print(f"  -> {len(df)} baris")

        except Exception as e:
            print(f"ERROR {kode_kab}: {e}")

        time.sleep(1)
    context.close()
    browser.close()

hasil = pd.concat(all_df, ignore_index=True)
# nama folder
folder = "scrap_anomali_usaha"
# buat folder jika belum ada
os.makedirs(folder, exist_ok=True)
# timestamp
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
# nama file
file_excel = os.path.join(folder, f"anomali_usaha_sumsel_{timestamp}.xlsx")
hasil.to_excel(file_excel, index=False)
print("Total baris:", len(hasil))
print("Selesai")
print(f"File tersimpan: {file_excel}")
