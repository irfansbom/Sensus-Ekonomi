"""
Scraper progres KBLI per kecamatan dari Dashboard SE2026.

Fitur:
- Tiap kecamatan yang berhasil diambil langsung disimpan ke file Excel
  terpisah (aman kalau proses berhenti di tengah jalan).
- Kalau dijalankan ulang, kecamatan yang sudah punya file akan otomatis
  dilewati (resume otomatis).
- Saat butuh captcha ulang, user bisa pilih: ENTER untuk lanjut,
  atau ketik 'd' untuk berhenti (data yang sudah diambil tetap aman).
"""

import os
import time
from datetime import datetime

import pandas as pd
from playwright.sync_api import sync_playwright
from playwright_stealth import Stealth
import threading

# ---------------------------------------------------------------------------
# Konfigurasi
# ---------------------------------------------------------------------------

URL_DASHBOARD = "https://dashboard-se2026.apps.bps.go.id"
URL_API = "https://dashboard-se2026.apps.bps.go.id/api/agregat/fasih"

SESSION_FILE = "Dashboard Scrapper/session_dash.json"
FOLDER_OUTPUT = "scrap_progres_kbli"
MAX_RETRY_PER_KEC = (
    5  # batas retry supaya tidak infinite loop kalau captcha terus gagal
)
DELAY_ANTAR_KEC = 3  # detik, jeda antar request supaya tidak dianggap bot

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


def buat_browser_context(playwright):
    """Buka browser + context dengan session yang sudah login sebelumnya."""
    browser = playwright.chromium.launch(
        headless=False,
        args=["--disable-blink-features=AutomationControlled"],
    )
    context = browser.new_context(
        storage_state=SESSION_FILE,
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
    return browser, context, page


def path_file_kec(kec: str) -> str:
    return os.path.join(FOLDER_OUTPUT, f"progres_kbli_{kec}.xlsx")


def input_with_timeout(prompt: str, timeout: float) -> str | None:
    """
    Sama seperti input(), tapi kalau tidak ada respon dalam `timeout` detik,
    return None (dianggap user tekan ENTER / lanjut).
    """
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
    """
    Ambil data satu kecamatan dari API.
    Return dict/list JSON kalau berhasil, None kalau gagal setelah retry habis.
    Raise StopScrapingException kalau user memilih berhenti.
    """
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

        if "application/json" in content_type:
            return response.json()

        print(f"  [{kec}] Response bukan JSON, sepertinya perlu captcha ulang.")
        print("  Cuplikan response:", response.text()[:300])

        page.goto(URL_DASHBOARD)
        time.sleep(70)

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


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main():
    os.makedirs(FOLDER_OUTPUT, exist_ok=True)

    kec_selesai = []
    kec_gagal = []
    stop_requested = False

    with sync_playwright() as p:
        browser, context, page = buat_browser_context(p)

        page.goto(URL_DASHBOARD) 
        # time.sleep(10)

        input("Selesaikan captcha dulu, lalu tekan ENTER...")

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


def gabungkan_semua_file():
    """
    Utilitas terpisah: gabungkan semua file per-kecamatan yang sudah
    tersimpan di FOLDER_OUTPUT menjadi satu file rekap.
    Panggil manual kapan saja setelah sebagian/semua kecamatan selesai.
    """
    import glob

    files = glob.glob(os.path.join(FOLDER_OUTPUT, "progres_kbli_*.xlsx"))
    if not files:
        print("Belum ada file kecamatan yang tersimpan.")
        return

    all_df = [pd.read_excel(f) for f in files]
    hasil = pd.concat(all_df, ignore_index=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_rekap = os.path.join(FOLDER_OUTPUT, f"rekap_gabungan_{timestamp}.xlsx")
    hasil.to_excel(file_rekap, index=False)

    print(f"Total file digabung : {len(files)}")
    print(f"Total baris gabungan: {len(hasil)}")
    print(f"File rekap tersimpan: {file_rekap}")


if __name__ == "__main__":
    main()
    # Setelah semua/sebagian kecamatan selesai, jalankan ini untuk gabungkan:
    # gabungkan_semua_file()
