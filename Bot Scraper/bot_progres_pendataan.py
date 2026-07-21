"""
Rekap Progres Pendataan Usaha & Keluarga
Pipeline: scrap dashboard (usaha & keluarga) -> pivot -> merge
          -> gabung dengan daftar PPL/PML -> export ke Excel & upsert ke SQLite.
"""

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

load_dotenv()
# ── Konfigurasi umum ─────────────────────────────────────────────────────

FORMAT_TANGGAL_JAM = "%d-%m-%Y_%H-%M-%S"
TANGGAL_JAM = datetime.now().strftime(FORMAT_TANGGAL_JAM)

URL_DASHBOARD = "https://dashboard-se2026.apps.bps.go.id"
URL_API = "https://dashboard-se2026.apps.bps.go.id/api/agregat/fasih"
SESSION_STATE = "Dashboard Scrapper/session_dash.json"
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

FILE_DAFTAR_PPL = Path("../20260608 Rekap Prelist_16_ppl_pml.xlsx")
FOLDER_OUTPUT_EXCEL = Path("../Rekap Progres Pendataan")
FOLDER_OUTPUT_DB = Path("../SQLLITE")
DB_PATH = FOLDER_OUTPUT_DB / "rekap_progress_pendataan.db"
TABLE_NAME = "rekap_progress_pendataan"

# Konfigurasi tiap jenis scraping: indikator yang diminta ke API,
# folder backup, dan nama file backup.
SCRAP_CONFIG = {
    "usaha": {
        "indikator": "108,109,110,10264,10265,10247,10266,10268,14,882,"
        "10691,10693,10694,10695,10696,10246",
        "folder_backup": Path("../scrap_status_usaha"),
        "prefix_file": "status_usaha_sls_sumsel",
    },
    "keluarga": {
        "indikator": "14,15,16,17,18,19,20,21,59",
        "folder_backup": Path("../scrap_status_keluarga"),
        "prefix_file": "status_keluarga_sls_sumsel",
    },
}

KOLOM_WILAYAH = ["id_wilayah", "kd_kab", "nama_kecamatan", "nama_desa", "nama_sls"]

KOLOM_INDIKATOR_USAHA = [
    "Jumlah UB Prelist Awal",
    "Jumlah UM Prelist Awal",
    "Jumlah UMK Prelist Awal",
    "Jumlah Usaha Ditemukan (BKU)",
    "Jumlah Usaha Ditutup (BKU)",
    "Jumlah Usaha Ganda (BKU)",
    "Jumlah Usaha Tidak Ditemukan (BKU)",
    "Jumlah Usaha Baru (BKU)",
    "Jumlah Keluarga Prelist Awal",
    "Jumlah Usaha Ditemukan (Usaha Keluarga)",
    "Jumlah Usaha Tutup (Usaha Keluarga)",
    "Jumlah Usaha Ganda (Usaha Keluarga)",
    "Jumlah Usaha Tidak Ditemukan (Usaha Keluarga)",
    "Jumlah Usaha Baru (Usaha Keluarga)",
    "Progres Pendataan Usaha dalam Keluarga",
]

KOLOM_INDIKATOR_KELUARGA = [
    "Jumlah Keluarga Prelist Awal_x",
    "Jumlah Keluarga Ditemukan",
    "Jumlah Keluarga Meninggal",
    "Jumlah Keluarga Tidak Eligible",
    "Jumlah Keluarga Tidak Dapat Ditemui Sampai Akhir Pendataan",
    "Jumlah Keluarga Tidak Ditemukan",
    "Jumlah Keluarga Baru",
    "Jumlah Keluarga Menolak Didata",
]

KOLOM_UNTUK_PRELIST_USAHA = [
    "Jumlah UB Prelist Awal",
    "Jumlah UM Prelist Awal",
    "Jumlah UMK Prelist Awal",
]

KOLOM_UNTUK_REALISASI_USAHA = [
    "Jumlah Usaha Ditemukan (BKU)",
    "Jumlah Usaha Baru (BKU)",
    "Jumlah Usaha Ditemukan (Usaha Keluarga)",
    "Jumlah Usaha Baru (Usaha Keluarga)",
]

KOLOM_UNTUK_REALISASI_KELUARGA = ["Jumlah Keluarga Ditemukan", "Jumlah Keluarga Baru"]

KOLOM_AKHIR = (
    ["id_wilayah", "kd_kab", "nama_sls"]
    + KOLOM_INDIKATOR_USAHA
    + [k for k in KOLOM_INDIKATOR_KELUARGA if k != "Jumlah Keluarga Prelist Awal_x"]
    + [
        "jumlah_prelist_usaha",
        "jumlah_usaha_realisasi",
        "jumlah_prelist_keluarga",
        "jumlah_keluarga_realisasi",
    ]
)


# ── Bagian 1: Scraping dashboard ────────────────────────────────────────
class BotTerdeteksi(Exception):
    """Dipicu kalau situs mendeteksi request ini sebagai bot. Sengaja dibuat
    berbeda dari error biasa supaya seluruh proses scraping BERHENTI TOTAL,
    bukan retry/lanjut diam-diam ke kabupaten lain."""
DELAY_ANTAR_KABUPATEN_DETIK = (120,180)  # jeda acak antara 15-30 detik


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
                params={"level": "sub_sls", "indikator": indikator, "kabupaten": kab},
                timeout=120000,
            )
        except Exception as e:
            print(f"  [{kab}] percobaan {percobaan} - Request error: {e}")
            time.sleep(5)
            continue
 
        content_type = response.headers.get("content-type", "")
        print(f"  [{kab}] percobaan {percobaan} - Status: {response.status} - "
              f"Content-Type: {content_type}")
 
        if response.status == 200 and "application/json" in content_type:
            return response.json()
 
        cuplikan = response.text()[:500]
 
        if "Bot Detected" in cuplikan or "terdeteksi sebagai bot" in cuplikan:
            print(f"  [{kab}] Kena deteksi bot. Menunggu 70 detik untuk cooldown...")
            time.sleep(70)
            page.goto(URL_DASHBOARD)
            continue  # coba lagi kabupaten yang sama, masih dalam batas MAX_RETRY_PER_SLS
 
        if response.status in (401, 403):
            print(f"  [{kab}] Status {response.status} - sesi login kemungkinan expired.")
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
            # storage_state=SESSION_STATE,
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


# ── Bagian 2: Transformasi (pivot & agregasi) ───────────────────────────


def pivot_indikator(df: pd.DataFrame, kolom_indikator: list[str]) -> pd.DataFrame:
    """Pivot data long -> wide berdasarkan nama_indikator, lalu pastikan semua
    kolom indikator ada (walau semua NaN di data asli) tanpa memicu cartesian
    product dari dropna=False."""
    df_pivot = df.pivot_table(
        index=KOLOM_WILAYAH,
        columns="nama_indikator",
        values="total_value",
        aggfunc="first",
    ).reset_index()

    for kolom in kolom_indikator:
        if kolom not in df_pivot.columns:
            df_pivot[kolom] = pd.NA

    df_pivot = df_pivot[KOLOM_WILAYAH + kolom_indikator]
    df_pivot[kolom_indikator] = df_pivot[kolom_indikator].fillna(0)

    return df_pivot.sort_values(by="id_wilayah").reset_index(drop=True)


def jumlahkan_kolom(
    df: pd.DataFrame, kolom_sumber: list[str], nama_kolom_baru: str
) -> pd.DataFrame:
    """Konversi kolom sumber ke numerik lalu jumlahkan jadi satu kolom baru."""
    for kolom in kolom_sumber:
        df[kolom] = pd.to_numeric(df[kolom], errors="coerce")
    df[nama_kolom_baru] = df[kolom_sumber].sum(axis=1)
    return df


def proses_data_usaha(df_usaha_mentah: pd.DataFrame) -> pd.DataFrame:
    df_pivot = pivot_indikator(df_usaha_mentah, KOLOM_INDIKATOR_USAHA)
    df_pivot = jumlahkan_kolom(
        df_pivot, KOLOM_UNTUK_PRELIST_USAHA, "jumlah_prelist_usaha"
    )
    df_pivot = jumlahkan_kolom(
        df_pivot, KOLOM_UNTUK_REALISASI_USAHA, "jumlah_usaha_realisasi"
    )
    return df_pivot


def proses_data_keluarga(df_keluarga_mentah: pd.DataFrame) -> pd.DataFrame:
    df_pivot = pivot_indikator(df_keluarga_mentah, KOLOM_INDIKATOR_KELUARGA)
    df_pivot = jumlahkan_kolom(
        df_pivot, KOLOM_UNTUK_REALISASI_KELUARGA, "jumlah_keluarga_realisasi"
    )
    return df_pivot


# ── Bagian 3: Merge & gabung dengan daftar PPL/PML ──────────────────────


def gabungkan_usaha_keluarga(
    df_usaha_pivot: pd.DataFrame, df_keluarga_pivot: pd.DataFrame
) -> pd.DataFrame:
    df_merged = df_usaha_pivot.merge(
        df_keluarga_pivot, on="id_wilayah", how="outer", suffixes=("_x", "_y")
    )
 
    # Kolom kd_kab & nama_sls muncul di kedua df; ambil yang tidak kosong.
    df_merged["nama_sls"] = df_merged["nama_sls_x"].combine_first(
        df_merged["nama_sls_y"]
    )
    df_merged["kd_kab"] = df_merged["kd_kab_x"].combine_first(df_merged["kd_kab_y"])
    df_merged = df_merged.rename(
        columns={"Jumlah Keluarga Prelist Awal": "jumlah_prelist_keluarga"}
    )

    return df_merged[KOLOM_AKHIR]


def gabungkan_dengan_ppl_pml(df_merged: pd.DataFrame) -> pd.DataFrame:
    df_merged = df_merged.copy()
    df_merged["id_wilayah"] = df_merged["id_wilayah"].astype(str).str.strip()
 
    df_ppl = pd.read_excel(FILE_DAFTAR_PPL, dtype=str)
    df_ppl = df_ppl.dropna(subset=["IDSUBSLS_25_2"]).reset_index(drop=True)
    df_ppl["IDSUBSLS_25_2"] = df_ppl["IDSUBSLS_25_2"].astype(str).str.strip()
 
    df_hasil = df_merged.merge(
        df_ppl[["IDSUBSLS_25_2", "PPL", "PML"]],
        left_on="id_wilayah",
        right_on="IDSUBSLS_25_2",
        how="left",
    ).drop(columns="IDSUBSLS_25_2")
 
    jumlah_tidak_match = df_hasil["PPL"].isna().sum()
    if jumlah_tidak_match > 0:
        print(
            f"Peringatan: {jumlah_tidak_match} dari {len(df_hasil)} baris tidak "
            "ketemu PPL/PML-nya. Kemungkinan ada perbedaan format id_wilayah "
            "(misal angka nol di depan hilang saat scraping). Cek contoh:"
        )
        print(df_hasil.loc[df_hasil["PPL"].isna(), "id_wilayah"].head(5).tolist())
        print("Bandingkan dengan contoh id di daftar PPL:")
        print(df_ppl["IDSUBSLS_25_2"].head(5).tolist())
 
    return df_hasil


# ── Bagian 4: Export ke Excel & SQLite ──────────────────────────────────


def simpan_ke_excel(df: pd.DataFrame) -> Path:
    FOLDER_OUTPUT_EXCEL.mkdir(parents=True, exist_ok=True)
    path_output = FOLDER_OUTPUT_EXCEL / f"rekap_progres_pendataan_{TANGGAL_JAM}.xlsx"
    df.to_excel(path_output, index=False)
    print("Excel tersimpan di:", path_output)
    return path_output


def pastikan_tabel_dan_kolom(df: pd.DataFrame, db_path: Path, table: str) -> None:
    """Buat tabel + unique index kalau belum ada; tambah kolom baru kalau df
    punya kolom yang belum dikenal tabel."""
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table,)
    )
    tabel_ada = cursor.fetchone()

    if not tabel_ada:
        df.head(0).to_sql(table, conn, if_exists="replace", index=False)
        cursor.execute(
            f'CREATE UNIQUE INDEX IF NOT EXISTS idx_id_wilayah ON "{table}" (id_wilayah)'
        )
    else:
        cursor.execute(f'PRAGMA table_info("{table}")')
        kolom_ada = {row[1] for row in cursor.fetchall()}
        for kolom in df.columns:
            if kolom not in kolom_ada:
                cursor.execute(f'ALTER TABLE "{table}" ADD COLUMN "{kolom}" TEXT')
                print(f"Kolom '{kolom}' ditambahkan ke tabel {table}")

    conn.commit()
    conn.close()


def upsert_ke_sqlite(
    df: pd.DataFrame, tanggal_jam: str, db_path: Path = DB_PATH, table: str = TABLE_NAME
) -> None:
    df = df.copy()
    df["last_update"] = tanggal_jam
 
    FOLDER_OUTPUT_DB.mkdir(parents=True, exist_ok=True)
    pastikan_tabel_dan_kolom(df, db_path, table)
 
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
 
    kolom_str = ", ".join(f'"{k}"' for k in df.columns)
    placeholder = ", ".join(["?"] * len(df.columns))
    data = df.where(pd.notnull(df), None).values.tolist()
 
    cursor.executemany(
        f'INSERT OR REPLACE INTO "{table}" ({kolom_str}) VALUES ({placeholder})', data
    )
    conn.commit()
 
    # Diagnostic: berapa baris di DB yang TIDAK ikut ter-upsert run ini,
    # supaya kelihatan jelas mana SLS yang last_update-nya tidak berubah
    # karena memang tidak ada di df run ini (bukan karena bug upsert).
    id_di_df = set(df["id_wilayah"].astype(str))
    cursor.execute(f'SELECT id_wilayah FROM "{table}"')
    id_di_db = {row[0] for row in cursor.fetchall()}
    id_tidak_tersentuh = id_di_db - id_di_df
 
    conn.close()
 
    print(f"{len(df)} baris berhasil di-upsert ke {table} pada {tanggal_jam}")
    if id_tidak_tersentuh:
        print(
            f"Catatan: {len(id_tidak_tersentuh)} id_wilayah di database TIDAK "
            "ikut ter-upsert run ini (last_update tetap yang lama) karena "
            "tidak ada di data hasil scraping run ini."
        )


# ── Main ──────────────────────────────────────────────────────────────────


def main() -> None:
    # df_usaha_mentah, df_keluarga_mentah = jalankan_scraping()
    df_usaha_mentah = pd.read_excel(
        "scrap_status_usaha\status_usaha_sls_sumsel_20260715_135259.xlsx"
    )
    df_keluarga_mentah = pd.read_excel(
        "scrap_status_keluarga\status_keluarga_sls_sumsel_20260715_141229.xlsx"
    )
    df_usaha_pivot = proses_data_usaha(df_usaha_mentah)
    df_keluarga_pivot = proses_data_keluarga(df_keluarga_mentah)
  
    df_merged = gabungkan_usaha_keluarga(df_usaha_pivot, df_keluarga_pivot)
    print(df_merged.columns)
    # df_final = gabungkan_dengan_ppl_pml(df_merged)

    # simpan_ke_excel(df_final)
    # upsert_ke_sqlite(df_final, TANGGAL_JAM)


if __name__ == "__main__":
    main()
