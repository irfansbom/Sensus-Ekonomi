import os
import sqlite3
import time
from datetime import datetime
from pathlib import Path
import random
import pandas as pd
from playwright.sync_api import sync_playwright
from playwright_stealth import Stealth
import threading
import pyotp
from dotenv import load_dotenv

# ---------------------------------------------------------------------------
# Konfigurasi
# ---------------------------------------------------------------------------

ENV_PATH = Path(__file__).resolve().parent / ".env"
load_dotenv(dotenv_path=ENV_PATH)

# ── Konfigurasi umum ─────────────────────────────────────────────────────

FORMAT_TANGGAL = "%d-%m-%Y"
FORMAT_TANGGAL_JAM = "%d-%m-%Y_%H-%M-%S"
TANGGAL_HARI_INI = datetime.now().strftime(FORMAT_TANGGAL)
TANGGAL_JAM = datetime.now().strftime(FORMAT_TANGGAL_JAM)

# Kredensial dari .env, sama seperti script rekap progres pendataan.
DASHBOARD_USERNAME = os.environ.get("DASHBOARD_USERNAME")
DASHBOARD_PASSWORD = os.environ.get("DASHBOARD_PASSWORD")
DASHBOARD_OTP_SECRET = os.environ.get("DASHBOARD_OTP_SECRET")

URL_DASHBOARD = "https://dashboard-se2026.apps.bps.go.id/se2026"
URL_API = "https://dashboard-se2026.apps.bps.go.id/api/agregat/fasih"

FOLDER_OUTPUT = "../scrap_progres_kbli"
MAX_RETRY_PER_KEC = (
    5 
)
DELAY_ANTAR_KEC = 30  

FOLDER_OUTPUT_REKAP = Path("../rekap_progres_kbli")

FOLDER_OUTPUT_DB = Path("../SQLLITE")
DB_PATH = FOLDER_OUTPUT_DB / "rekap_progres_kbli.db"
TABLE_NAME = "rekap_progres_kbli"


INDIKATOR_JUMLAH_USAHA = [
    'Jumlah Usaha Kategori KBLI "A"',
    'Jumlah Usaha Kategori KBLI "B"',
    "Jumlah Usaha C",
    "Jumlah Usaha D",
    "Jumlah Usaha E",
    "Jumlah Usaha F",
    "Jumlah Usaha G",
    "Jumlah Usaha H",
    "Jumlah Usaha I",
    "Jumlah Usaha J",
    "Jumlah Usaha K",
    'Jumlah Usaha Kategori KBLI "L" (Aktivitas Keuangan dan Asuransi)',
    "Jumlah Usaha M",
    "Jumlah Usaha N, O",
    "Jumlah Usaha P",
    "Jumlah Usaha Q",
    "Jumlah Usaha R",
    "Jumlah Usaha S, U",
    "Jumlah Usaha T",
    "Jumlah Usaha V",
]

INDIKATOR_NILAI_TAMBAH = [
    'Nilai Tambah Kategori KBLI "A"',
    "Nilai Tambah B",
    "Nilai Tambah C",
    "Nilai Tambah D",
    "Nilai Tambah E",
    "Nilai Tambah F",
    "Nilai Tambah G",
    "Nilai Tambah H",
    "Nilai Tambah I",
    "Nilai Tambah J",
    "Nilai Tambah K",
    'Nilai Tambah Kategori KBLI "L" (Aktivitas Keuangan dan Asuransi)',
    "Nilai Tambah M",
    "Nilai Tambah N, O",
    "Nilai Tambah P",
    "Nilai Tambah Q",
    "Nilai Tambah R",
    "Nilai Tambah S, U",
    "Nilai Tambah T",
    "Nilai Tambah V",
]

INDIKATOR_TOTAL_OMSET = [
    'Total Omzet Kategori KBLI "A"',
    "Total Omzet B",
    "Total Omzet C",
    "Total Omzet D",
    "Total Omzet E",
    "Total Omzet F",
    "Total Omzet G",
    "Total Omzet H",
    "Total Omzet I",
    "Total Omzet J",
    "Total Omzet K",
    'Total Omzet Kategori KBLI "L" (Aktivitas Keuangan dan Asuransi)',
    "Total Omzet M",
    "Total Omzet N, O",
    "Total Omzet P",
    "Total Omzet Q",
    "Total Omzet R",
    "Total Omzet S, U",
    "Total Omzet T",
    "Total Omzet V",
]

INDIKATOR_TOTAL_OUTPUT = [
    "Total Output A",
    "Total Output B",
    "Total Output C",
    "Total Output D",
    "Total Output E",
    "Total Output F",
    "Total Output G",
    "Total Output H",
    "Total Output I",
    "Total Output J",
    "Total Output K",
    "Total Output L",
    "Total Output M",
    "Total Output N,O",
    "Total Output P",
    "Total Output Q",
    "Total Output R",
    "Total Output S,U",
    "Total Output T",
    "Total Output V",
]

INDIKATOR_TENAGA_KERJA = [
    "Total Tenaga Kerja A",
    "Total Tenaga Kerja B",
    "Total Tenaga Kerja C",
    "Total Tenaga Kerja D",
    "Total Tenaga Kerja E",
    "Total Tenaga Kerja F",
    "Total Tenaga Kerja G",
    "Total Tenaga Kerja H",
    "Total Tenaga Kerja I",
    "Total Tenaga Kerja J",
    "Total Tenaga Kerja K",
    "Total Tenaga Kerja L",
    "Total Tenaga Kerja M",
    "Total Tenaga Kerja N,O",
    "Total Tenaga Kerja P",
    "Total Tenaga Kerja Q",
    "Total Tenaga Kerja R",
    "Total Tenaga Kerja S,U",
    "Total Tenaga Kerja T",
    "Total Tenaga Kerja V",
]

INDIKATOR_TOTAL_UPAH = [
    "Total Upah dan Gaji A",
    "Total Upah dan Gaji B",
    "Total Upah dan Gaji C",
    "Total Upah dan Gaji D",
    "Total Upah dan Gaji E",
    "Total Upah dan Gaji F",
    "Total Upah dan Gaji G",
    "Total Upah dan Gaji H",
    "Total Upah dan Gaji I",
    "Total Upah dan Gaji J",
    "Total Upah dan Gaji K",
    "Total Upah dan Gaji L",
    "Total Upah dan Gaji M",
    "Total Upah dan Gaji N,O",
    "Total Upah dan Gaji P",
    "Total Upah dan Gaji Q",
    "Total Upah dan Gaji R",
    "Total Upah dan Gaji S,U",
    "Total Upah dan Gaji T",
    "Total Upah dan Gaji V",
]

KOLOM_AKHIR = [
    "id_wilayah",
    "nama_wilayah",
    "jumlah_usaha",
    "nilai_tambah",
    "total_omset",
    "total_output",
    "tenaga_kerja",
    "total_upah",
]

INDIKATOR = (
    "60,61,62,11519,11520,11521,63,64,65,11522,11523,11524,66,67,68,11525,11526,11527,"
    "69,70,71,11528,11529,11530,72,73,74,11531,11532,11533,75,76,77,11534,11535,11536,"
    "78,79,80,11537,11538,11539,81,82,83,11540,11541,11542,84,85,86,11543,11544,11545,"
    "87,88,89,11546,11547,11548,90,91,92,11549,11550,11551,93,94,95,11552,11553,11554,"
    "96,97,98,11555,11556,11557,10254,10255,10256,11558,11559,11560,99,160,161,11561,"
    "11562,11563,162,100,163,11564,11565,11566,164,165,101,11567,11568,11569,10260,10262,"
    "10263,11570,11571,11572,11509,11510,11511,11573,11574,11575,11512,11513,11514,11576,"
    "11577,11578"
)

KODE_KEC_LIST = [
    "1601052",
    "1601070",
    "1601080",
    "1601081",
    "1601082",
    "1601083",
    "1601090",
    "1601091",
    "1601092",
    "1601093",
    "1601130",
    "1601131",
    "1601140",
    "1602010",
    "1602011",
    "1602020",
    "1602021",
    "1602022",
    "1602023",
    "1602030",
    "1602031",
    "1602040",
    "1602041",
    "1602050",
    "1602051",
    "1602060",
    "1602120",
    "1602121",
    "1602130",
    "1602131",
    "1602140",
    "1603010",
    "1603011",
    "1603012",
    "1603020",
    "1603021",
    "1603031",
    "1603032",
    "1603033",
    "1603040",
    "1603050",
    "1603051",
    "1603060",
    "1603061",
    "1603062",
    "1603070",
    "1603071",
    "1603090",
    "1603091",
    "1603092",
    "1603093",
    "1603094",
    "1603095",
    "1604011",
    "1604012",
    "1604040",
    "1604041",
    "1604042",
    "1604043",
    "1604050",
    "1604051",
    "1604052",
    "1604060",
    "1604061",
    "1604062",
    "1604063",
    "1604111",
    "1604112",
    "1604113",
    "1604114",
    "1604120",
    "1604121",
    "1604122",
    "1604123",
    "1604131",
    "1604132",
    "1604133",
    "1605030",
    "1605031",
    "1605032",
    "1605040",
    "1605041",
    "1605050",
    "1605051",
    "1605060",
    "1605061",
    "1605070",
    "1605071",
    "1605072",
    "1605080",
    "1605090",
    "1606010",
    "1606020",
    "1606021",
    "1606022",
    "1606023",
    "1606030",
    "1606031",
    "1606040",
    "1606041",
    "1606090",
    "1606091",
    "1606092",
    "1606100",
    "1606101",
    "1606102",
    "1607010",
    "1607020",
    "1607021",
    "1607030",
    "1607031",
    "1607032",
    "1607040",
    "1607041",
    "1607050",
    "1607051",
    "1607060",
    "1607061",
    "1607070",
    "1607080",
    "1607081",
    "1607090",
    "1607091",
    "1607100",
    "1607101",
    "1607110",
    "1607111",
    "1608010",
    "1608020",
    "1608021",
    "1608022",
    "1608030",
    "1608040",
    "1608041",
    "1608050",
    "1608051",
    "1608060",
    "1608061",
    "1608070",
    "1608071",
    "1608080",
    "1608090",
    "1608091",
    "1608100",
    "1608101",
    "1608102",
    "1609010",
    "1609011",
    "1609012",
    "1609020",
    "1609030",
    "1609031",
    "1609032",
    "1609040",
    "1609041",
    "1609050",
    "1609051",
    "1609060",
    "1609061",
    "1609070",
    "1609080",
    "1609081",
    "1609090",
    "1609091",
    "1609100",
    "1609101",
    "1610010",
    "1610011",
    "1610012",
    "1610020",
    "1610021",
    "1610030",
    "1610031",
    "1610040",
    "1610041",
    "1610042",
    "1610050",
    "1610051",
    "1610052",
    "1610060",
    "1610061",
    "1610062",
    "1611010",
    "1611020",
    "1611030",
    "1611031",
    "1611040",
    "1611050",
    "1611051",
    "1611060",
    "1611070",
    "1611071",
    "1612010",
    "1612020",
    "1612030",
    "1612040",
    "1612050",
    "1613010",
    "1613020",
    "1613030",
    "1613040",
    "1613050",
    "1613060",
    "1613070",
    "1671010",
    "1671011",
    "1671020",
    "1671021",
    "1671022",
    "1671030",
    "1671031",
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
    "1672010",
    "1672020",
    "1672021",
    "1672030",
    "1672031",
    "1672040",
    "1673010",
    "1673011",
    "1673020",
    "1673030",
    "1673040",
    "1674011",
    "1674012",
    "1674021",
    "1674022",
    "1674031",
    "1674032",
    "1674041",
    "1674042",
]


class StopScrapingException(Exception):
    """Dilempar saat user memilih untuk berhenti di tengah proses."""

    pass


# ---------------------------------------------------------------------------
# Fungsi bantu
# ---------------------------------------------------------------------------


SELECTOR_USERNAME = "xpath=//*[@id='username']"
SELECTOR_PASSWORD = "xpath=//*[@id='password']"
SELECTOR_OTP = "xpath=//*[@id='otp']"
SELECTOR_TOMBOL_LOGIN = "xpath=//*[@id='v-0']/button"

def isi_username_password(page) -> None:
    page.wait_for_selector(SELECTOR_USERNAME, state="visible", timeout=15000)
    page.fill(SELECTOR_USERNAME, DASHBOARD_USERNAME)
    page.fill(SELECTOR_PASSWORD, DASHBOARD_PASSWORD)
    page.press(SELECTOR_PASSWORD, "Enter")
    time.sleep(5)
    print("Username & password terisi, form disubmit.")


def isi_otp(page, timeout_ms: int = 20000) -> None:
    page.wait_for_selector(SELECTOR_OTP, state="visible", timeout=timeout_ms)

    totp = pyotp.TOTP(DASHBOARD_OTP_SECRET)
    otp_code = totp.now()

    page.fill(SELECTOR_OTP, otp_code)
    page.press(SELECTOR_OTP, "Enter")

    time.sleep(5)
    print(f"OTP ({otp_code}) terisi, form disubmit.")


def klik_tombol_login(page, max_percobaan: int = 5, timeout_ms: int = 15000) -> bool:
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
            time.sleep(random.uniform(1, 5))
            page.locator(SELECTOR_TOMBOL_LOGIN).click(timeout=5000)
        except Exception as e:
            print(f"  Percobaan {percobaan} - klik gagal: {e}")
            page.wait_for_timeout(1500)
            continue

        page.wait_for_timeout(2000)

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


def goto_dashboard_aman(page, percobaan: int = 3) -> bool:
    """Navigasi ke URL_DASHBOARD dengan lebih toleran terhadap koneksi
    lambat: pakai wait_until='domcontentloaded' (tidak perlu nunggu semua
    resource/gambar selesai load) + timeout lebih panjang, dan retry
    beberapa kali kalau timeout. Return True kalau berhasil, False kalau
    tetap gagal setelah semua percobaan (proses tetap lanjut, tidak crash)."""
    for i in range(1, percobaan + 1):
        try:
            page.goto(URL_DASHBOARD, timeout=60_000, wait_until="domcontentloaded")
            return True
        except Exception as e:
            print(f"  goto dashboard percobaan {i} gagal: {e}")
            time.sleep(3)
    print("  Gagal membuka dashboard setelah beberapa percobaan, lanjut tanpa reload.")
    return False


def login_dashboard(page) -> None:
    """Bungkus alur login lengkap: klik tombol login -> (captcha manual kalau
    perlu) -> isi username/password/OTP."""
    goto_dashboard_aman(page)
    login_selesai_otomatis = klik_tombol_login(page)

    if not login_selesai_otomatis:
        if page.locator(SELECTOR_USERNAME).count() > 0:
            isi_username_password(page)
            isi_otp(page)


def buat_browser_context(playwright):
    """Buka browser + context dengan session yang sudah login sebelumnya."""
    browser = playwright.chromium.launch(
        headless=False,
        args=["--disable-blink-features=AutomationControlled"],
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
    login_dashboard(page)
    return browser, context, page


def path_file_kec(kec: str) -> str:
    return os.path.join(FOLDER_OUTPUT, f"progres_kbli_{kec}.xlsx")


def input_with_timeout(prompt: str, timeout: float) -> str | None :
    result = {}

    def _get_input():
        try:
            result["value"] = input(prompt)
        except Exception:
            result["value"] = None

    thread = threading.Thread(target=_get_input, daemon=True)
    thread.start()
    thread.join(timeout)

    if thread.is_alive():
        # Timeout habis, user belum jawab -> anggap lanjut
        print("\n  Waktu habis (60 detik), otomatis lanjut...")
        return None  # None = dianggap lanjut, bukan 'd'

    return result.get("value")


def fetch_api(context, page, kec: str):
    for percobaan in range(1, MAX_RETRY_PER_KEC + 1):
        try:
            response = context.request.get(
                URL_API,
                params={
                    "level": "sub_sls",
                    "indikator": INDIKATOR,
                    "kecamatan": kec,
                },
                timeout=120_000,
            )
        except Exception as e:
            print(f"  [{kec}] percobaan {percobaan} - Request error: {e}")
            time.sleep(5)
            continue

        content_type = response.headers.get("content-type", "")
        print(
            f"  [{kec}] percobaan {percobaan} "
            f"- Status: {response.status} - Content-Type: {content_type}"
        )

        if response.status == 200 and "application/json" in content_type:
            return response.json()

        if "application/json" in content_type and response.status in (401, 403):
            # Body-nya JSON tapi status 401/403 -> sesi login expired,
            # bukan captcha. Cukup buka ulang dashboard & tunggu sebentar,
            # lalu coba lagi kecamatan yang sama (masih dalam batas retry).
            print(
                f"  [{kec}] Status {response.status} - sesi login kemungkinan "
                "expired. Reload dashboard & tunggu 5 detik..."
            )
            goto_dashboard_aman(page)
            time.sleep(5)
            continue

        cuplikan = response.text()[:500]

        if "Bot Detected" in cuplikan or "terdeteksi sebagai bot" in cuplikan:
            print(f"  [{kec}] Kena deteksi bot. Menunggu 70 detik untuk cooldown...")
            time.sleep(120)
            goto_dashboard_aman(page)
            continue  # coba lagi kecamatan yang sama, masih dalam batas MAX_RETRY_PER_KEC

        if "text/html" in content_type :
            print(
                f"  [{kec}] Status {response.status} - sesi login kemungkinan "
                "expired. Reload dashboard & tunggu 5 detik..."
            )
            goto_dashboard_aman(page)
            time.sleep(5)
            continue

        print(f"  [{kec}] Response bukan JSON, sepertinya perlu captcha ulang.")
        print("  Cuplikan response:", cuplikan[:300])

        goto_dashboard_aman(page)
        time.sleep(120)

        pilihan = input_with_timeout(
            f"  Selesaikan captcha untuk kec {kec}.\n"
            f"  Tekan ENTER untuk lanjut, atau ketik 'd' lalu ENTER untuk "
            f"STOP (data yang sudah ada akan tetap aman) [auto-lanjut dalam 60 detik]: ",
            timeout=60,
        )
        if pilihan is not None and pilihan.strip().lower() == "d":
            raise StopScrapingException()

    print(f"  [{kec}] GAGAL setelah {MAX_RETRY_PER_KEC} percobaan, dilewati.")
    return None


def proses_satu_kecamatan(context, page, kec: str) -> bool:
    """
    Ambil data satu kecamatan dan simpan ke file Excel sendiri.
    Return True kalau berhasil disimpan, False kalau tidak ada data.
    """
    data = fetch_api(context, page, kec)
    if not data:
        return False

    df = pd.DataFrame(data)
    df["kode_kecamatan"] = kec

    file_kec = path_file_kec(kec)
    df.to_excel(file_kec, index=False)
    print(f"  -> {len(df)} baris tersimpan ke {file_kec}")
    return True


def jalankan_scraping_anomali() -> tuple[pd.DataFrame, pd.DataFrame]:
    with sync_playwright() as p:
            browser, context, page = buat_browser_context(p)
            # Tidak perlu page.goto(URL_DASHBOARD) lagi di sini -- login_dashboard()
            # yang dipanggil di dalam buat_browser_context() sudah goto + login.
            # Navigasi kedua yang redundant ini penyebab TimeoutError kalau
            # dashboard lambat merender ulang setelah submit OTP.
            time.sleep(5)
            kec_selesai = []
            kec_gagal = []
            for kec in KODE_KEC_LIST:
                print(f"Proses Kec {kec}")
    
                # resume otomatis: skip kalau file kec ini sudah pernah berhasil
                if os.path.exists(path_file_kec(kec)):
                    print(f"  -> Sudah ada file untuk kec {kec}, dilewati.")
                    continue
    
                try:
                    berhasil = proses_satu_kecamatan(context, page, kec)
                    if berhasil:
                        kec_selesai.append(kec)
                    else:
                        kec_gagal.append(kec)
    
                except StopScrapingException:
                    idx = KODE_KEC_LIST.index(kec)
                    sisa = KODE_KEC_LIST[idx:]
                    print(f"\nDihentikan oleh user pada kec {kec}.")
                    print(f"Sisa {len(sisa)} kecamatan belum diproses:")
                    print(sisa)
                    stop_requested = True
                    break
    
                except Exception as e:
                    print(f"ERROR Kec {kec}: {e}")
                    kec_gagal.append(kec)
    
                time.sleep(DELAY_ANTAR_KEC)
    
            context.close()
            browser.close()
    return kec_selesai, kec_gagal


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main():
    os.makedirs(FOLDER_OUTPUT, exist_ok=True)

    kec_selesai, kec_gagal = jalankan_scraping_anomali()
    stop_requested = False

    print("\n" + "=" * 50)
    print(f"Kecamatan berhasil diambil sesi ini : {len(kec_selesai)}")
    print(f"Kecamatan gagal/tidak ada data      : {len(kec_gagal)}")
    if kec_gagal:
        print("Daftar kecamatan gagal:", kec_gagal)
    if stop_requested:
        print(
            "Status: DIHENTIKAN MANUAL - jalankan ulang script untuk lanjut otomatis."
        )
    else:
        print("Status: SELESAI semua kecamatan dalam daftar.")
    print("=" * 50)


# ---------------------------------------------------------------------------
# Pivot, agregasi & upsert ke SQLite
# ---------------------------------------------------------------------------


def _jumlahkan_satu_kategori(
    df: pd.DataFrame, daftar_indikator: list[str], nama_kolom_baru: str
) -> pd.DataFrame:
    """Filter baris sesuai daftar_indikator, lalu jumlahkan total_value per
    SLS (id_wilayah + nama_wilayah). Hasil: df 3 kolom (id_wilayah,
    nama_wilayah, nama_kolom_baru)."""
    subset = df[df["nama_indikator"].isin(daftar_indikator)]
    return (
        subset.groupby(["id_wilayah", "nama_wilayah"], as_index=False)["total_value"]
        .sum()
        .rename(columns={"total_value": nama_kolom_baru})
    )


def pivot_dan_agregasi(df_mentah: pd.DataFrame) -> pd.DataFrame:
    """Ubah data long (satu baris per SLS x indikator) jadi satu baris per
    SLS dengan 6 kolom agregat (jumlah_usaha, nilai_tambah, dst), mengikuti
    pola yang sama dengan rekap_progres_kbli.ipynb."""
    kategori = [
        (INDIKATOR_JUMLAH_USAHA, "jumlah_usaha"),
        (INDIKATOR_NILAI_TAMBAH, "nilai_tambah"),
        (INDIKATOR_TOTAL_OMSET, "total_omset"),
        (INDIKATOR_TOTAL_OUTPUT, "total_output"),
        (INDIKATOR_TENAGA_KERJA, "tenaga_kerja"),
        (INDIKATOR_TOTAL_UPAH, "total_upah"),
    ]

    df_gabung = None
    for daftar_indikator, nama_kolom in kategori:
        df_kategori = _jumlahkan_satu_kategori(df_mentah, daftar_indikator, nama_kolom)
        df_gabung = (
            df_kategori
            if df_gabung is None
            else df_gabung.merge(
                df_kategori, on=["id_wilayah", "nama_wilayah"], how="outer"
            )
        )

    return df_gabung[KOLOM_AKHIR]


def konversi_kolom_numerik_ke_int(df: pd.DataFrame) -> pd.DataFrame:
    """Pastikan kolom-kolom agregat tersimpan sebagai integer (bukan float),
    supaya tidak muncul '.0' di SQLite. Pakai Int64 (nullable) supaya nilai
    NaN tetap aman (tidak error) kalau memang ada baris yang datanya kosong."""
    df = df.copy()
    kolom_numerik = [k for k in KOLOM_AKHIR if k not in ("id_wilayah", "nama_wilayah")]
    for kolom in kolom_numerik:
        df[kolom] = pd.to_numeric(df[kolom], errors="coerce").round().astype("Int64")
    return df


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
    """Upsert (INSERT OR REPLACE) berdasarkan id_wilayah sebagai unique key,
    persis pola yang sama dengan bot_progres_pendataan.py."""
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


def gabungkan_semua_file():
    """
    Utilitas terpisah: gabungkan semua file per-kecamatan yang sudah
    tersimpan di FOLDER_OUTPUT menjadi satu file rekap, lalu pivot+agregasi
    dan upsert ke SQLite. Panggil manual kapan saja setelah sebagian/semua
    kecamatan selesai.
    """
    import glob

    files = glob.glob(os.path.join(FOLDER_OUTPUT, "progres_kbli_*.xlsx"))
    if not files:
        print("Belum ada file kecamatan yang tersimpan.")
        return

    all_df = [pd.read_excel(f) for f in files]
    hasil = pd.concat(all_df, ignore_index=True)

    tanggal_jam = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    FOLDER_OUTPUT_REKAP.mkdir(parents=True, exist_ok=True)
    file_rekap = FOLDER_OUTPUT_REKAP / f"rekap_progres_kbli_{timestamp}.xlsx"
    hasil.to_excel(file_rekap, index=False)

    print(f"Total file digabung : {len(files)}")
    print(f"Total baris gabungan: {len(hasil)}")
    print(f"File rekap tersimpan: {file_rekap}")

    print("Proses Pivot & Agregasi")
    df_pivot = pivot_dan_agregasi(hasil)
    df_pivot = konversi_kolom_numerik_ke_int(df_pivot)

    print("Proses Upsert ke SQLite")
    upsert_ke_sqlite(df_pivot, tanggal_jam)


if __name__ == "__main__":
    main()
    # Setelah semua/sebagian kecamatan selesai, jalankan ini untuk gabungkan
    # + upsert ke SQLite:
    gabungkan_semua_file()
