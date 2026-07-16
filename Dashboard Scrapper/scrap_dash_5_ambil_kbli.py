from playwright.sync_api import sync_playwright
import pandas as pd
import os
import time
from datetime import datetime


kode_kec_list = [
    "1671010",
    "1671011",
    "1671020",
    "1671021",
    "1671022",
    "1671030",
    "1671040",
    "1671041",
    "1671050",
    "1671051",
    "1671060",
    "1671061",
    "1671062",
    "1671070",
    "1671071",
    "1671080",
    "1671081",
]

all_df = []

with sync_playwright() as p:

    browser = p.chromium.launch(headless=False)

    context = browser.new_context(storage_state="scrap_dashboard/session_dash.json")
    page = context.new_page()

    page.goto("https://dashboard-se2026.apps.bps.go.id")
    input("Selesaikan captcha dulu, lalu tekan ENTER...")
    for kode_kec in kode_kec_list:

        print(f"Proses {kode_kec}")

        try:

            # https://dashboard-se2026.apps.bps.go.id/api/agregat/fasih?level=desa&indikator=61,64,67,70,73,76,79,82,85,88,91,94,97,10255,160,100,165,10262&kecamatan=1671010
            response = context.request.get(
                "https://dashboard-se2026.apps.bps.go.id/api/agregat/fasih",
                params={
                    "level": "desa",
                    "kecamatan": kode_kec,
                    "indikator": "61,64,67,70,73,76,79,82,85,88,91,94,97,10255,160,100,165,10262"
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
            df["kode_kecamatan"] = kode_kec

            all_df.append(df)

            print(f"  -> {len(df)} baris")

        except Exception as e:
            print(f"ERROR {kode_kec}: {e}")

        time.sleep(1)
    context.close()
    browser.close()

hasil = pd.concat(all_df, ignore_index=True)
# nama folder
folder = "scrap_usaha_perkbli"
# buat folder jika belum ada
os.makedirs(folder, exist_ok=True)
# timestamp
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
# nama file
file_excel = os.path.join(folder, f"usaha_perkbli_palembang_{timestamp}.xlsx")
hasil.to_excel(file_excel, index=False)
print("Total baris:", len(hasil))
print("Selesai")
print(f"File tersimpan: {file_excel}")
