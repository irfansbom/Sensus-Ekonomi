from playwright.sync_api import sync_playwright
import pandas as pd
import os
import time
from datetime import datetime
from playwright_stealth import Stealth
import json

kode_kab_list = [
    "1601",
    "1602",
    "1603",
    "1604",
    "1605",
    "1606",
    "1607",
    "1608",
    "1609",
    "1610",
    "1611",
    "1612",
    "1613",
    "1671",
    "1672",
    "1673",
    "1674",
]


all_df = []

def fetch_api(page, kode_kab):
    url = (
        "https://dashboard-se2026.apps.bps.go.id/api/mikro/anomali-case-kab"
        f"?kode_kabupaten={kode_kab}"
        "&indikator=136,137,139,140,141,142,144"
        "&sudah_indikator=47,48,50,51,52,53"
        "&type=keluarga"
        "&anomali_no=1,2,3,4,5,6,7"
    )
    result = page.evaluate(
        """async (url) => {
            const res = await fetch(url, { credentials: 'include' });
            const text = await res.text();
            return { status: res.status, contentType: res.headers.get('content-type'), body: text };
        }""",
        url,
    )
    return result


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
            # response = context.request.get(
            #     "https://dashboard-se2026.apps.bps.go.id/api/mikro/anomali-case-kab",
            #     params={
            #         "kode_kabupaten": kode_kab,
            #         "indikator": "136,137,139,140,141,142,144",
            #         "sudah_indikator": "47,48,50,51,52,53",
            #         "type": "keluarga",
            #         "anomali_no": "1,2,3,4,5,6,7",
            #     },
            # )
            result = fetch_api(page, kode_kab)

            print("Status:", result["status"])
            print("Content-Type:", result["contentType"])

            if not result["contentType"] or "application/json" not in result["contentType"]:
                print(result["body"][:1000])
                continue

            data = json.loads(result["body"])

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
folder = "scrap_anomali_keluarga"
# buat folder jika belum ada
os.makedirs(folder, exist_ok=True)
# timestamp
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
# nama file
file_excel = os.path.join(folder, f"anomali_keluarga_sumsel_{timestamp}.xlsx")
hasil.to_excel(file_excel, index=False)
print("Total baris:", len(hasil))
print("Selesai")
print(f"File tersimpan: {file_excel}")
