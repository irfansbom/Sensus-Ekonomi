from playwright.sync_api import sync_playwright
import pandas as pd
import os
import time
from datetime import datetime
from playwright_stealth import Stealth

URL_DASHBOARD = "https://dashboard-se2026.apps.bps.go.id"
URL_API = "https://dashboard-se2026.apps.bps.go.id/api/agregat/fasih"
FOLDER_OUTPUT = "scrap_status_keluarga"
MAX_RETRY_PER_SLS = (5)

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

def minta_data_sls(context, page, kab: str):
    for percobaan in range(1, MAX_RETRY_PER_SLS + 1):
        try:
            response = context.request.get(
                URL_API,
                params={
                    "level": "sub_sls",
                    "indikator": "14,15,16,17,18,19,20,21,59",
                    "kabupaten": kab,
                },
                timeout=120000,  # naikkan jadi 60 detik
            )
        except Exception as e:
            print(f"  [{kab}] percobaan {percobaan} - Request error: {e}")
            time.sleep(5)
            continue  # coba lagi tanpa langsung nyerah ke kabupaten berikutnya

        content_type = response.headers.get("content-type", "")
        print(
            f"  [{kab}] percobaan {percobaan} - Status: {response.status} - Content-Type: {content_type}"
        )

        if "application/json" in content_type:
            return response.json()

        print(f"  [{kab}] Response bukan JSON, sepertinya perlu captcha ulang.")
        print("  Cuplikan response:", response.text()[:300])

        page.goto(URL_DASHBOARD)
        input(f"  Selesaikan captcha untuk lanjut scrap Kab {kab}, lalu tekan ENTER...")

    print(f"  [{kab}] GAGAL setelah {MAX_RETRY_PER_SLS} percobaan, dilewati.")
    return None


def main():
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

        page.goto(URL_DASHBOARD)
        input("Selesaikan captcha dulu, lalu tekan ENTER...")

        for kab in kode_kab_list:
            print(f"Proses Kab {kab}")
            try:
                data = minta_data_sls(context, page, kab)
                if data is None:
                    continue

                df = pd.DataFrame(data)
                df["kd_kab"] = kab
                all_df.append(df)
                print(f"  -> {len(df)} baris")

            except Exception as e:
                print(f"ERROR Kab {kab}: {e}")

            time.sleep(3)

        context.close()
        browser.close()

    if not all_df:
        print("Tidak ada data yang berhasil diambil.")
        return

    hasil = pd.concat(all_df, ignore_index=True)
    os.makedirs(FOLDER_OUTPUT, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_excel = os.path.join(
        FOLDER_OUTPUT, f"status_keluarga_sls_sumsel_{timestamp}.xlsx"
    )
    hasil.to_excel(file_excel, index=False)

    print("Total baris:", len(hasil))
    print("Selesai")
    print(f"File tersimpan: {file_excel}")


if __name__ == "__main__":
    main()
