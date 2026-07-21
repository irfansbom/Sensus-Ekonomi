import re
import sqlite3
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
from playwright.sync_api import sync_playwright
from playwright_stealth import Stealth
import pyotp
import os
from dotenv import load_dotenv
import random


FORMAT_TANGGAL_JAM = "%d-%m-%Y_%H-%M-%S"
TANGGAL_JAM = datetime.now().strftime(FORMAT_TANGGAL_JAM)

URL_DASHBOARD = "https://dashboard-se2026.apps.bps.go.id"
URL_API = "https://dashboard-se2026.apps.bps.go.id/api/mikro/anomali-case-kab"

#   "https://dashboard-se2026.apps.bps.go.id/api/mikro/anomali-case-kab",
#                 params={
#                     "kode_kabupaten": kode_kab,
#                     "indikator": "128,129,130,131,132,133,134,135",
#                     "sudah_indikator": "40,41,42,43,44,45,46,46",
#                     "type": "usaha",
#                     "anomali_no": "1,2,3,4,5,6,7,8",
#                 },


MAX_RETRY_PER_SLS = 5

DASHBOARD_USERNAME = os.environ.get("DASHBOARD_USERNAME")
DASHBOARD_PASSWORD = os.environ.get("DASHBOARD_PASSWORD")
DASHBOARD_OTP_SECRET = os.environ.get("DASHBOARD_OTP_SECRET")

SELECTOR_USERNAME = "xpath=//*[@id='username']"
SELECTOR_PASSWORD = "xpath=//*[@id='password']"
SELECTOR_OTP = "xpath=//*[@id='otp']"


KODE_KAB_LIST = [
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

class BotTerdeteksi(Exception):
    """Dipicu kalau situs mendeteksi request ini sebagai bot. Sengaja dibuat
    berbeda dari error biasa supaya seluruh proses scraping BERHENTI TOTAL,
    bukan retry/lanjut diam-diam ke kabupaten lain."""


DELAY_ANTAR_KABUPATEN_DETIK = (15, 30)  # jeda acak antara 15-30 detik


def minta_data_sls(context, page, kab: str, indikator: str) -> list | None:
    """Minta data satu kabupaten ke API. Retry otomatis kalau:
    - request error (koneksi putus, timeout, dll)
    - status bukan 200 (misal 401 = sesi login expired, 400 = bad request)
    - kena deteksi bot dari situs (tunggu 70 detik cooldown, lalu coba lagi)

    Semua retry ini dibatasi oleh MAX_RETRY_PER_SLS -- kalau habis, kabupaten
    ini dilewati dan lanjut ke kabupaten berikutnya (bukan loop tanpa batas).
    """
    for percobaan in range(1, MAX_RETRY_PER_SLS + 1):
        try:
            response = context.request.get(
                URL_API,
                 params={
                    "kode_kabupaten": kab,
                    "indikator": "128,129,130,131,132,133,134,135",
                    "sudah_indikator": "40,41,42,43,44,45,46,46",
                    "type": "usaha",
                    "anomali_no": "1,2,3,4,5,6,7,8",
                },
                timeout=120000,
            )
        except Exception as e:
            print(f"  [{kab}] percobaan {percobaan} - Request error: {e}")
            time.sleep(5)
            continue

        content_type = response.headers.get("content-type", "")
        print(
            f"  [{kab}] percobaan {percobaan} - Status: {response.status} - "
            f"Content-Type: {content_type}"
        )

        if response.status == 200 and "application/json" in content_type:
            return response.json()

        cuplikan = response.text()[:500]

        if "Bot Detected" in cuplikan or "terdeteksi sebagai bot" in cuplikan:
            print(f"  [{kab}] Kena deteksi bot. Menunggu 70 detik untuk cooldown...")
            time.sleep(70)
            page.goto(URL_DASHBOARD)
            continue  # coba lagi kabupaten yang sama, masih dalam batas MAX_RETRY_PER_SLS

        if response.status in (401, 403):
            print(
                f"  [{kab}] Status {response.status} - sesi login kemungkinan expired."
            )
            print("  Cuplikan response:", cuplikan[:300])
            time.sleep(5)
            continue

        if response.status >= 400:
            print(f"  [{kab}] Status {response.status} - request gagal, coba lagi...")
            print("  Cuplikan response:", cuplikan[:300])
            time.sleep(5)
            continue

        print(f"  [{kab}] Response bukan JSON, sepertinya perlu captcha ulang.")
        print("  Cuplikan response:", cuplikan[:300])
        page.goto(URL_DASHBOARD)
        # input(f"  Selesaikan captcha untuk lanjut scrap Kab {kab}, lalu tekan ENTER...")

    print(f"  [{kab}] GAGAL setelah {MAX_RETRY_PER_SLS} percobaan, dilewati.")
    return None


def scrap_satu_jenis(context, page, indikator: str) -> pd.DataFrame:
    """Scrap semua kabupaten untuk satu jenis indikator (usaha atau keluarga).
    Kalau satu kabupaten gagal terus setelah semua retry (termasuk cooldown
    deteksi bot), kabupaten itu dilewati dan proses lanjut ke kabupaten
    berikutnya."""
    semua_df = []

    for kab in KODE_KAB_LIST:
        print(f"Proses Kab {kab}")
        data = minta_data_sls(context, page, kab, indikator)
        if data is not None:
            df = pd.DataFrame(data)
            df["kd_kab"] = kab
            semua_df.append(df)
            print(f"  -> {len(df)} baris")

        jeda = random.uniform(*DELAY_ANTAR_KABUPATEN_DETIK)
        print(f"  Jeda {jeda:.1f} detik sebelum kabupaten berikutnya...")
        time.sleep(jeda)

    if not semua_df:
        raise RuntimeError("Tidak ada data yang berhasil diambil untuk indikator ini.")

    return pd.concat(semua_df, ignore_index=True)


def simpan_backup_excel(df: pd.DataFrame, folder: Path, prefix: str) -> Path:
    """Simpan hasil scraping mentah sebagai backup, sebelum diproses lebih lanjut."""
    folder.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path_output = folder / f"{prefix}_{timestamp}.xlsx"
    df.to_excel(path_output, index=False)
    print("Backup tersimpan di:", path_output)
    return path_output


SELECTOR_TOMBOL_LOGIN = "xpath=//*[@id='v-0']/button"


def isi_username_password(page) -> None:
    """Isi form username & password, lalu submit dengan menekan Enter."""
    page.wait_for_selector(SELECTOR_USERNAME, state="visible", timeout=15000)
    page.fill(SELECTOR_USERNAME, DASHBOARD_USERNAME)
    page.fill(SELECTOR_PASSWORD, DASHBOARD_PASSWORD)
    page.press(SELECTOR_PASSWORD, "Enter")
    print("Username & password terisi, form disubmit.")


def isi_otp(page, timeout_ms: int = 20000) -> None:
    """Isi kode OTP dari TOTP generator, lalu submit dengan Enter.
    Kode OTP baru di-generate saat elemen sudah muncul, supaya tidak
    keburu expired kalau ada delay sebelumnya."""
    page.wait_for_selector(SELECTOR_OTP, state="visible", timeout=timeout_ms)

    totp = pyotp.TOTP(DASHBOARD_OTP_SECRET)
    otp_code = totp.now()

    page.fill(SELECTOR_OTP, otp_code)
    page.press(SELECTOR_OTP, "Enter")
    print(f"OTP ({otp_code}) terisi, form disubmit.")


def klik_tombol_login(page, max_percobaan: int = 5, timeout_ms: int = 15000) -> bool:
    """Klik tombol login, verifikasi ada perubahan nyata di halaman.
    Kalau form username langsung muncul (tanpa captcha), sekalian isi
    username/password/OTP di sini.

    Return True kalau proses login (username+password+OTP) sudah selesai
    ditangani otomatis di dalam fungsi ini, False kalau belum (misal karena
    masih perlu captcha manual dulu, atau tombol login tidak ditemukan sama
    sekali karena sudah login lewat session_state)."""
    try:
        page.wait_for_selector(
            SELECTOR_TOMBOL_LOGIN, timeout=timeout_ms, state="visible"
        )
    except Exception as e:
        print(f"Tombol login tidak ditemukan (mungkin sudah login): {e}")
        return False

    url_sebelum = page.url

    for percobaan in range(1, max_percobaan + 1):
        try:
            page.locator(SELECTOR_TOMBOL_LOGIN).click(timeout=5000)
        except Exception as e:
            print(f"  Percobaan {percobaan} - klik gagal: {e}")
            page.wait_for_timeout(1500)
            continue

        page.wait_for_timeout(2000)  # beri waktu JS merespons klik

        url_sesudah = page.url
        ada_perubahan_url = url_sesudah != url_sebelum
        ada_captcha = (
            page.locator("iframe[src*='captcha'], .captcha, #captcha").count() > 0
        )
        ada_form_username = page.locator(SELECTOR_USERNAME).count() > 0

        if ada_perubahan_url or ada_captcha or ada_form_username:
            print(
                f"Tombol login berhasil diklik (percobaan {percobaan}), halaman berubah."
            )

            if ada_form_username:
                isi_username_password(page)
                isi_otp(page)
                time.sleep(5)
                return True

            print("Form username belum muncul (kemungkinan perlu captcha manual dulu).")
            return False

        print(
            f"  Percobaan {percobaan} - klik terkirim tapi belum ada perubahan, coba lagi..."
        )

    print(
        "Tombol login diklik tapi tidak terdeteksi ada perubahan setelah beberapa percobaan."
    )
    return False


def jalankan_scraping() -> tuple[pd.DataFrame, pd.DataFrame]:
    """Scrap data usaha & keluarga dalam satu sesi browser, lalu backup masing-masing
    ke Excel. Mengembalikan (df_usaha, df_keluarga) mentah (belum di-pivot)."""
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True, args=["--disable-blink-features=AutomationControlled"]
        )
        context = browser.new_context(
            accept_downloads=True,
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1366, "height": 768},
            locale="id-ID",
            timezone_id="Asia/Jakarta",
        )
        page = context.new_page()
        Stealth().apply_stealth_sync(page)

        page.goto(URL_DASHBOARD)
        klik_tombol_login(page)

        # input("Selesaikan captcha dulu, lalu tekan ENTER...")
        # input
        print("=== Scraping data usaha ===")
        df_usaha = scrap_satu_jenis(context, page, SCRAP_CONFIG["usaha"]["indikator"])

        print("=== Scraping data keluarga ===")
        df_keluarga = scrap_satu_jenis(
            context, page, SCRAP_CONFIG["keluarga"]["indikator"]
        )

        context.close()
        browser.close()

    simpan_backup_excel(
        df_usaha,
        SCRAP_CONFIG["usaha"]["folder_backup"],
        SCRAP_CONFIG["usaha"]["prefix_file"],
    )
    simpan_backup_excel(
        df_keluarga,
        SCRAP_CONFIG["keluarga"]["folder_backup"],
        SCRAP_CONFIG["keluarga"]["prefix_file"],
    )

    return df_usaha, df_keluarga


def main() -> None:
    df_usaha_mentah, df_keluarga_mentah = jalankan_scraping()
    # df_usaha_mentah = pd.read_excel("../scrap_status_usaha/status_usaha_sls_sumsel_20260716_201826.xlsx")
    # df_keluarga_mentah = pd.read_excel("../scrap_status_keluarga/status_keluarga_sls_sumsel_20260716_202033.xlsx")
    df_usaha_pivot = proses_data_usaha(df_usaha_mentah)
    df_keluarga_pivot = proses_data_keluarga(df_keluarga_mentah)

    df_merged = gabungkan_usaha_keluarga(df_usaha_pivot, df_keluarga_pivot)
    df_final = gabungkan_dengan_ppl_pml(df_merged)

    simpan_ke_excel(df_final)
    upsert_ke_sqlite(df_final, TANGGAL_JAM)


if __name__ == "__main__":
    main()
