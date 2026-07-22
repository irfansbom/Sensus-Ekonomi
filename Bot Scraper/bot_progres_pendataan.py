"""
Rekap Progres Pendataan Usaha & Keluarga
Pipeline: scrap dashboard (usaha & keluarga) -> pivot -> merge
          -> gabung dengan daftar PPL/PML -> export ke Excel & upsert ke SQLite.
"""
 
import os
import random
import re
import sqlite3
import time
from datetime import datetime
from pathlib import Path
 
import pandas as pd
import pyotp
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
from playwright_stealth import Stealth
 

# Cari .env di folder yang sama dengan file script ini, terlepas dari dari
# direktori mana script dijalankan (penting karena beda OS/terminal punya
# working directory yang beda-beda).
ENV_PATH = Path(__file__).resolve().parent / ".env"
load_dotenv(dotenv_path=ENV_PATH)
 
# ── Konfigurasi umum ─────────────────────────────────────────────────────
 
FORMAT_TANGGAL_JAM = "%d-%m-%Y_%H-%M-%S"
TANGGAL_JAM = datetime.now().strftime(FORMAT_TANGGAL_JAM)
 
URL_DASHBOARD = "https://dashboard-se2026.apps.bps.go.id"
URL_API = "https://dashboard-se2026.apps.bps.go.id/api/agregat/fasih"
# SESSION_STATE = "Dashboard Scrapper/session_dash.json"
MAX_RETRY_PER_SLS = 5

# Kredensial login diambil dari file .env (lihat .env.example), tidak
# di-hardcode di sini supaya aman kalau file ini ikut ter-commit ke Git.
DASHBOARD_USERNAME = os.environ.get("DASHBOARD_USERNAME")
DASHBOARD_PASSWORD = os.environ.get("DASHBOARD_PASSWORD")
DASHBOARD_OTP_SECRET = os.environ.get("DASHBOARD_OTP_SECRET")
 
_KREDENSIAL = {
    "DASHBOARD_USERNAME": DASHBOARD_USERNAME,
    "DASHBOARD_PASSWORD": DASHBOARD_PASSWORD,
    "DASHBOARD_OTP_SECRET": DASHBOARD_OTP_SECRET,
}
_KOSONG = [k for k, v in _KREDENSIAL.items() if not v]
 
if _KOSONG:
    raise RuntimeError(
        f"Variabel berikut kosong/tidak ditemukan: {', '.join(_KOSONG)}.\n"
        f"Dicari dari file: {ENV_PATH} (ada: {ENV_PATH.exists()}).\n"
        "Pastikan file .env ada persis di folder yang sama dengan script ini "
        "dan berisi baris seperti:\n"
        "  DASHBOARD_USERNAME=muh.prayitno\n"
        "  DASHBOARD_PASSWORD=Prayitno_26\n"
        "  DASHBOARD_OTP_SECRET=MRVGI53EMIYTKMTHOZ3W64RXJU2G6R3H"
    )
 
SELECTOR_USERNAME = "xpath=//*[@id='username']"
SELECTOR_PASSWORD = "xpath=//*[@id='password']"
SELECTOR_OTP = "xpath=//*[@id='otp']"
 
KODE_KAB_LIST = [
    "1601", "1602", "1603", "1604", "1605", "1606", "1607", "1608", "1609",
    "1610", "1611", "1612", "1613", "1671", "1672", "1673", "1674",
]
 
FILE_DAFTAR_PPL = Path("../20260608 Rekap Prelist_16_ppl_pml.xlsx")
FOLDER_OUTPUT_EXCEL = Path("../Rekap Progres Pendataan")
FOLDER_OUTPUT_DB = Path("../SQLLITE")
DB_PATH = FOLDER_OUTPUT_DB / "rekap_progress_pendataan.db"
TABLE_NAME = "rekap_progress_pendataan"
 
# Semua indikator (usaha + keluarga) diminta dalam SATU request per kabupaten,
# supaya tidak perlu merge dua sumber data lagi -- ini yang jadi sumber bug
# berulang (suffix _x/_y, kolom duplikat, dll).
SCRAP_CONFIG = {
    "indikator": "108,109,110,10264,10265,10247,10266,10268,14,882,10691,"
                 "10693,10694,10695,10696,10246,15,16,17,18,19,20,21,59",
    "folder_backup": Path("../scrap_status_pendataan"),
    "prefix_file": "status_pendataan_sls_sumsel",
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
    "Jumlah Usaha Ditemukan (Usaha Keluarga)",
    "Jumlah Usaha Tutup (Usaha Keluarga)",
    "Jumlah Usaha Ganda (Usaha Keluarga)",
    "Jumlah Usaha Tidak Ditemukan (Usaha Keluarga)",
    "Jumlah Usaha Baru (Usaha Keluarga)",
]
 
KOLOM_INDIKATOR_KELUARGA = [
    "Jumlah Keluarga Prelist Awal",
    "Jumlah Keluarga Ditemukan",
    "Jumlah Keluarga Baru",
    "Jumlah Keluarga Meninggal",
    "Jumlah Keluarga Tidak Eligible",
    "Jumlah Keluarga Tidak Dapat Ditemui Sampai Akhir Pendataan",
    "Jumlah Keluarga Tidak Ditemukan",
]
 
# Gabungan kedua daftar indikator, dedupe tapi urutan tetap terjaga
# ("Jumlah Keluarga Prelist Awal" ada di kedua daftar, jadi cuma dihitung 1x).
KOLOM_INDIKATOR_GABUNGAN = list(
    dict.fromkeys(KOLOM_INDIKATOR_USAHA + KOLOM_INDIKATOR_KELUARGA)
)
 
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
 
# "Jumlah Keluarga Prelist Awal" diwakili oleh jumlah_prelist_keluarga (hasil
# rename), jadi versi mentahnya tidak perlu dobel muncul di output akhir.
KOLOM_AKHIR = (
    ["id_wilayah", "kd_kab", "nama_sls"]
    + [k for k in KOLOM_INDIKATOR_GABUNGAN if k != "Jumlah Keluarga Prelist Awal"]
    + [
        "jumlah_prelist_usaha",
        "jumlah_usaha_realisasi",
        "jumlah_prelist_keluarga",
        "jumlah_keluarga_realisasi",
    ]
)
 

# ── Bagian 1: Scraping dashboard ────────────────────────────────────────
 
# Jeda antar-request. Dibuat jauh lebih lambat daripada sebelumnya supaya
# pola trafik tidak terlihat seperti bot. Sesuaikan lagi kalau masih kena
# deteksi.
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
            print(f"  [{kab}] Kena deteksi bot. Menunggu 120 detik untuk cooldown...")
            time.sleep(120)
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
        input(f"  Selesaikan captcha untuk lanjut scrap Kab {kab}, lalu tekan ENTER...")
 
    print(f"  [{kab}] GAGAL setelah {MAX_RETRY_PER_SLS} percobaan, dilewati.")
    return None
 
 
def scrap_satu_jenis(context, page, indikator: str) -> pd.DataFrame:
    """Scrap semua kabupaten untuk satu jenis indikator (usaha atau keluarga).
    Kalau satu kabupaten gagal terus setelah semua retry (termasuk cooldown
    deteksi bot), kabupaten itu dilewati dan proses lanjut ke kabupaten
    berikutnya -- artinya SEMUA SLS di kabupaten itu tidak akan ter-upsert
    di run ini, dan last_update-nya di database tetap yang lama."""
    semua_df = []
    kab_gagal = []
 
    for kab in KODE_KAB_LIST:
        print(f"Proses Kab {kab}")
        data = minta_data_sls(context, page, kab, indikator)
        if data is not None:
            df = pd.DataFrame(data)
            df["kd_kab"] = kab
            semua_df.append(df)
            print(f"  -> {len(df)} baris")
        else:
            kab_gagal.append(kab)
 
        jeda = random.uniform(*DELAY_ANTAR_KABUPATEN_DETIK)
        print(f"  Jeda {jeda:.1f} detik sebelum kabupaten berikutnya...")
        time.sleep(jeda)
 
    if kab_gagal:
        print(
            f"\nPERINGATAN: {len(kab_gagal)} kabupaten GAGAL di-scrap dan dilewati: "
            f"{kab_gagal}\nSemua SLS di kabupaten ini TIDAK ter-upsert run ini "
            "(last_update di database tetap yang lama untuk SLS tersebut).\n"
        )
 
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
        page.wait_for_selector(SELECTOR_TOMBOL_LOGIN, timeout=timeout_ms, state="visible")
    except Exception as e:
        print(f"Tombol login tidak ditemukan (mungkin sudah login): {e}")
        return False
 
    url_sebelum = page.url
 
    for percobaan in range(1, max_percobaan + 1):
        try:
            time.sleep(random.uniform(1, 5))  # jeda acak sebelum klik
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
            print(f"Tombol login berhasil diklik (percobaan {percobaan}), halaman berubah.")
 
            if ada_form_username:
                isi_username_password(page)
                isi_otp(page)
                return True
 
            print("Form username belum muncul (kemungkinan perlu captcha manual dulu).")
            return False
 
        print(f"  Percobaan {percobaan} - klik terkirim tapi belum ada perubahan, coba lagi...")
 
    print("Tombol login diklik tapi tidak terdeteksi ada perubahan setelah beberapa percobaan.")
    return False
 
 
def jalankan_scraping() -> pd.DataFrame:
    """Scrap semua indikator (usaha & keluarga sekaligus) dalam satu sesi
    browser, lalu backup ke Excel. Mengembalikan df mentah (belum di-pivot)."""
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
        login_selesai_otomatis = klik_tombol_login(page)
 
        # Kalau login belum ditangani otomatis (misal karena captcha sempat
        # muncul duluan), selesaikan captcha manual dulu, baru lanjutkan
        # mengisi username/password/OTP secara otomatis.
        if not login_selesai_otomatis:
            input("Selesaikan captcha dulu, lalu tekan ENTER...")
            if page.locator(SELECTOR_USERNAME).count() > 0:
                isi_username_password(page)
                isi_otp(page)
 
        print("=== Scraping data pendataan (usaha & keluarga) ===")
        df_mentah = scrap_satu_jenis(context, page, SCRAP_CONFIG["indikator"])
 
        context.close()
        browser.close()
 
    simpan_backup_excel(
        df_mentah, SCRAP_CONFIG["folder_backup"], SCRAP_CONFIG["prefix_file"]
    )
 
    return df_mentah
 
 
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
 
 
def proses_data(df_mentah: pd.DataFrame) -> pd.DataFrame:
    """Pivot semua indikator (usaha & keluarga) sekaligus, lalu hitung
    kolom-kolom agregat. Tidak perlu merge dua df lagi karena datanya
    sudah satu sumber sejak awal."""
    df_pivot = pivot_indikator(df_mentah, KOLOM_INDIKATOR_GABUNGAN)
 
    df_pivot = jumlahkan_kolom(df_pivot, KOLOM_UNTUK_PRELIST_USAHA, "jumlah_prelist_usaha")
    df_pivot = jumlahkan_kolom(df_pivot, KOLOM_UNTUK_REALISASI_USAHA, "jumlah_usaha_realisasi")
    df_pivot = jumlahkan_kolom(
        df_pivot, KOLOM_UNTUK_REALISASI_KELUARGA, "jumlah_keluarga_realisasi"
    )
 
    df_pivot["jumlah_prelist_keluarga"] = pd.to_numeric(
        df_pivot["Jumlah Keluarga Prelist Awal"], errors="coerce"
    )
 
    return df_pivot[KOLOM_AKHIR]
 
 
# ── Bagian 3: Gabung dengan daftar PPL/PML ──────────────────────────────
 
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
 
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table,))
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
 
  
def konversi_kolom_numerik_ke_int(df: pd.DataFrame) -> pd.DataFrame:
    """Pastikan kolom-kolom indikator & agregat tersimpan sebagai integer
    (bukan float), supaya tidak muncul '.0' di Excel maupun SQLite.
    Pakai Int64 (nullable) dari pandas, bukan int biasa, supaya nilai NaN
    tetap aman (tidak error) kalau memang ada baris yang datanya kosong."""
    df = df.copy()
    kolom_numerik = [
        k for k in KOLOM_AKHIR
        if k not in ("id_wilayah", "kd_kab", "nama_sls")
    ]
    for kolom in kolom_numerik:
        df[kolom] = pd.to_numeric(df[kolom], errors="coerce").round().astype("Int64")
    return df
 
 
 
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
    print("Proses Scraping Dashboard")
    df_mentah = jalankan_scraping()
    print("Proses Pivot Data")
    df_pivot = proses_data(df_mentah)
    print("Proses Menggabung Petugas")
    df_final = gabungkan_dengan_ppl_pml(df_pivot)
    df_final = konversi_kolom_numerik_ke_int(df_final)
 
    simpan_ke_excel(df_final)
    upsert_ke_sqlite(df_final, TANGGAL_JAM)
 
if __name__ == "__main__":
    main()