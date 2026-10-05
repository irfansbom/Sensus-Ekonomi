import csv
import json
import os
import random
import time
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
from playwright_stealth import Stealth

# ============================================================
# KONFIGURASI
# ============================================================

ENV_PATH = Path(__file__).resolve().parent / ".env"
load_dotenv(dotenv_path=ENV_PATH)


# ------------------------------------------------------------
# URL
# ------------------------------------------------------------

LINK_BASE_URL = "https://fasih-dashboard.bps.go.id/"
LINK_SQLLAB_URL = "https://fasih-dashboard.bps.go.id/superset/sqllab/"
URL_SQL_EXECUTE = "https://fasih-dashboard.bps.go.id/api/v1/sqllab/execute/"


# ------------------------------------------------------------
# LOGIN
# ------------------------------------------------------------

SELECTOR_USERNAME = "xpath=//*[@id='username']"
SELECTOR_PASSWORD = "xpath=//*[@id='password']"
SELECTOR_TOMBOL_LOGIN = "xpath=//*[@id='container']/div[2]/form/button"

SQLLAB_USERNAME = os.environ.get("SQLLAB_USERNAME")
SQLLAB_PASSWORD = os.environ.get("SQLLAB_PASSWORD")


# ------------------------------------------------------------
# SCRAPING
# ------------------------------------------------------------

# Jumlah row yang diminta pada SQL
QUERY_LIMIT = 9000

# Jumlah retry jika request gagal
MAX_RETRY = 5

# Timeout request dalam milidetik
REQUEST_TIMEOUT = 120000


# ------------------------------------------------------------
# SQL LAB
# ------------------------------------------------------------

DATABASE_ID = 25

SQL_EDITOR_ID = "1036293"

# Tetap 1000 sesuai request browser SQL Lab
QUERY_LIMIT_SUPERSET = 1000

SCHEMA = "tgr_fd68e454"

TAB_NAME = "List Usaha dan Tagging"


# ------------------------------------------------------------
# OUTPUT
# ------------------------------------------------------------

OUTPUT_DIR = Path(__file__).resolve().parent / "hasil_scraping"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# ============================================================
# SQL DASAR
# ============================================================

SQL_BASE = """
SELECT
    assignment_id,
    assignment_status_alias,
    level_6_full_code,
    nama_usaha_bang,
    nama_kk,
    ada_keluarga_label,
    ada_bang_usaha_label,
    geotag_accuracy,
    geotag_latitude,
    geotag_longitude
FROM
    tgr_fd68e454.root_table
WHERE
    (
        ada_bang_usaha_value IN (1,2)
        OR
        ada_keluarga_value IN (1,2)
    )
"""


# ============================================================
# MEMBUAT SQL LIMIT / OFFSET
# ============================================================


def buat_sql(offset):
    """
    Membuat SQL berdasarkan OFFSET.

    Contoh:

        LIMIT 9000 OFFSET 0
        LIMIT 9000 OFFSET 9000
        LIMIT 9000 OFFSET 18000
    """

    return f"""
{SQL_BASE.strip()}

LIMIT {QUERY_LIMIT}
OFFSET {offset}
""".strip()


# ============================================================
# LOGIN
# ============================================================


def isi_username_password(page):
    """
    Mengisi username dan password kemudian submit.
    """
    page.wait_for_selector(SELECTOR_USERNAME, state="visible", timeout=15000)
    page.fill(SELECTOR_USERNAME, SQLLAB_USERNAME)
    page.fill(SELECTOR_PASSWORD, SQLLAB_PASSWORD)
    page.press(SELECTOR_PASSWORD, "Enter")
    print("Username & password terisi, " "form disubmit.")


def klik_tombol_login(page, max_percobaan=5, timeout_ms=15000):
    """
    Mendeteksi dan menekan tombol login.
    Jika form login tidak ditemukan,
    diasumsikan browser sudah login.
    """
    try:
        page.wait_for_selector(
            SELECTOR_TOMBOL_LOGIN, timeout=timeout_ms, state="visible"
        )
    except Exception:
        print("Tombol login tidak ditemukan.")
        print(
            "Kemungkinan browser sudah login "
            "atau membutuhkan proses autentikasi tambahan."
        )
        return False

    url_sebelum = page.url
    for percobaan in range(1, max_percobaan + 1):
        try:
            time.sleep(random.uniform(1, 3))
            page.locator(SELECTOR_TOMBOL_LOGIN).click(timeout=5000)
        except Exception as e:
            print(f"Percobaan {percobaan} " f"- klik login gagal: {e}")
            page.wait_for_timeout(1500)
            continue

        page.wait_for_timeout(2000)
        url_sesudah = page.url
        ada_perubahan_url = url_sesudah != url_sebelum
        ada_form_username = page.locator(SELECTOR_USERNAME).count() > 0

        if ada_perubahan_url or ada_form_username:
            print(f"Tombol login berhasil diproses " f"(percobaan {percobaan}).")
            if ada_form_username:
                isi_username_password(page)
            return True
        print(f"Percobaan {percobaan} " "belum ada perubahan, " "mencoba lagi...")

    return False


def login_dashboard(page):
    """
    Membuka dashboard dan melakukan login.
    """
    print("\n" + "=" * 70)
    print("LOGIN FASIH DASHBOARD")
    print("=" * 70)
    page.goto(LINK_BASE_URL, wait_until="domcontentloaded")
    login_selesai = klik_tombol_login(page)
    if not login_selesai:
        print("\nSilakan selesaikan login/captcha " "jika diperlukan.")
        input("Setelah login berhasil, " "tekan ENTER untuk melanjutkan...")
        if page.locator(SELECTOR_USERNAME).count() > 0:
            isi_username_password(page)
    page.wait_for_timeout(3000)
    print("\nLogin selesai.")
    print("URL:", page.url)


# ============================================================
# MENANGKAP REQUEST SQL LAB
# ============================================================


def pasang_penangkap_csrf(page):
    """
    Menangkap CSRF token dari request SQL Lab
    yang benar-benar dibuat oleh browser.
    Endpoint /security/csrf_token/ tidak digunakan
    karena server FASIH mengembalikan HTTP 403.
    """
    request_info = {"csrf_token": None, "headers": None, "payload": None}
    def handle_request(request):
        if "/api/v1/sqllab/execute/" not in request.url:
            return
        csrf_token = request.headers.get("x-csrftoken")
        if not csrf_token:
            return
        request_info["csrf_token"] = csrf_token
        request_info["headers"] = request.headers
        try:
            if request.post_data:
                request_info["payload"] = json.loads(request.post_data)
        except Exception:
            request_info["payload"] = None
        print("\n" + "=" * 70)
        print("REQUEST SQL LAB TERDETEKSI")
        print("=" * 70)
        print("Method :", request.method)
        print("URL    :", request.url)
        print("CSRF   : ADA")
        print("Referer:", request.headers.get("referer"))
        print("Content-Type:", request.headers.get("content-type"))
        print("\n[OK] CSRF token berhasil ditangkap.")
    page.on("request", handle_request)
    return request_info


# ============================================================
# MENUNGGU CSRF DARI REQUEST BROWSER
# ============================================================


def tunggu_csrf(page, request_info, timeout_detik=300):
    """
    Menunggu browser melakukan request SQL Lab
    dan menangkap CSRF token.

    User cukup menjalankan satu query manual
    dari SQL Lab.
    """
    print("\n" + "=" * 70)
    print("MENUNGGU REQUEST SQL LAB")
    print("=" * 70)
    print("""
        SQL Lab sudah dibuka.

        Silakan lakukan SATU kali query dari SQL Lab
        secara manual dengan tombol RUN.

        Gunakan query yang biasanya berhasil.

        Setelah request berhasil dikirim,
        program akan menangkap X-CSRFToken secara otomatis.

        JANGAN kirim atau tampilkan nilai token/cookie.
        """)

    batas_waktu = time.time() + timeout_detik

    while time.time() < batas_waktu:
        if request_info["csrf_token"]:
            print("\n[OK] CSRF token sudah tersedia.")
            return request_info["csrf_token"]
        page.wait_for_timeout(500)
    raise Exception(
        "Timeout menunggu request SQL Lab. "
        "Pastikan tombol RUN sudah ditekan "
        "dan query berhasil dijalankan."
    )


# ============================================================
# PAYLOAD SQL LAB
# ============================================================


def buat_payload(sql):
    """
    Membuat payload sesuai request browser SQL Lab.
    client_id tidak digunakan karena opsional.
    queryLimit tetap 1000.
    LIMIT SQL dikontrol oleh QUERY_LIMIT.
    """

    return {
        "database_id": DATABASE_ID,
        "json": True,
        "runAsync": False,
        "schema": SCHEMA,
        "sql": sql,
        "sql_editor_id": SQL_EDITOR_ID,
        "tab": TAB_NAME,
        "tmp_table_name": "",
        "select_as_cta": False,
        "ctas_method": "TABLE",
        "queryLimit": QUERY_LIMIT,
        "expand_data": True,
    }


# ============================================================
# EKSEKUSI SQL
# ============================================================


def execute_query(context, sql, csrf_token, offset):
    """
    Menjalankan SQL menggunakan session browser
    yang sudah login.
    """

    payload = buat_payload(sql)

    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "X-CSRFToken": csrf_token,
        "Referer": LINK_SQLLAB_URL,
    }

    for percobaan in range(1, MAX_RETRY + 1):
        print("\n" + "-" * 70)
        print(
            f"OFFSET {offset:,} | "
            f"LIMIT {QUERY_LIMIT:,} | "
            f"Percobaan {percobaan}/{MAX_RETRY}"
        )
        try:
            response = context.request.post(
                URL_SQL_EXECUTE, data=payload, headers=headers, timeout=REQUEST_TIMEOUT
            )
            print("HTTP Status:", response.status)
            content_type = response.headers.get("content-type", "")
            print("Content-Type:", content_type)
            response_text = response.text()
            # ------------------------------------------------
            # BERHASIL
            # ------------------------------------------------
            if response.status == 200:
                try:
                    result = response.json()
                    print("[OK] Query berhasil.")
                    return result
                except Exception as e:
                    print("Response berhasil " "tetapi bukan JSON:", e)
                    print(response_text[:2000])
                    return None

            # ------------------------------------------------
            # CSRF INVALID
            # ------------------------------------------------

            if response.status == 400:
                print("\nHTTP 400")
                print(response_text[:2000])
                if "CSRF token" in response_text or "csrf" in response_text.lower():
                    print("\n[WARNING] " "CSRF token tidak valid.")
                    return "REFRESH_CSRF"

            # ------------------------------------------------
            # SESSION
            # ------------------------------------------------
            elif response.status in (401, 403):
                print("\nSession kemungkinan " "sudah expired.")
                print(response_text[:1000])
                return None

            # ------------------------------------------------
            # ERROR LAIN
            # ------------------------------------------------

            else:
                print("\nRequest gagal.")
                print(response_text[:2000])
        except Exception as e:
            print(f"\nRequest error: {e}")

        # ----------------------------------------------------
        # RETRY
        # ----------------------------------------------------

        if percobaan < MAX_RETRY:
            waktu_tunggu = min(5 * percobaan, 30)
            print(f"Menunggu " f"{waktu_tunggu} detik...")
            time.sleep(waktu_tunggu)
    print(f"\n[GAGAL] OFFSET {offset:,}")
    return None


# ============================================================
# EKSTRAK ROW DARI RESPONSE
# ============================================================
def ambil_rows(result):
    """
    Mengambil data row dari beberapa kemungkinan
    struktur response Superset.
    """
    if result is None:
        return []
    if not isinstance(result, dict):
        return []
    # --------------------------------------------------------
    # Bentuk:
    #
    # {
    #     "data": [...]
    # }
    # --------------------------------------------------------
    data = result.get("data")
    if isinstance(data, list):
        return data

    # --------------------------------------------------------
    # Bentuk:
    #
    # {
    #     "result": [...]
    # }
    # --------------------------------------------------------

    result_data = result.get("result")

    if isinstance(result_data, list):

        for item in result_data:

            if not isinstance(item, dict):

                continue

            data = item.get("data")

            if isinstance(data, list):

                return data

    # --------------------------------------------------------
    # Bentuk:
    #
    # {
    #     "result": {
    #         "data": [...]
    #     }
    # }
    # --------------------------------------------------------

    if isinstance(result_data, dict):

        data = result_data.get("data")

        if isinstance(data, list):

            return data

    return []


# ============================================================
# STATE RESUME
# ============================================================


def buat_state_baru(nama_file):
    """
    Membuat state awal scraping.
    """

    return {
        "nama_file": str(nama_file),
        "next_offset": 0,
        "total_data": 0,
        "batch_ke": 0,
        "header_sudah_ditulis": False,
        "file_size": 0,
        "updated_at": datetime.now().isoformat(),
    }


def baca_state(state_file):
    """
    Membaca state scraping.
    """

    if not state_file.exists():

        return None

    try:

        with open(state_file, "r", encoding="utf-8") as f:

            return json.load(f)

    except Exception as e:

        print(f"[WARNING] " f"State tidak dapat dibaca: {e}")

        return None


def simpan_state(state_file, state):
    """
    Menyimpan state secara atomic.

    Menggunakan file temporary lalu replace,
    sehingga risiko state JSON rusak lebih kecil
    jika program berhenti tiba-tiba.
    """

    temp_file = state_file.with_suffix(".tmp")

    state["updated_at"] = datetime.now().isoformat()

    with open(temp_file, "w", encoding="utf-8") as f:

        json.dump(state, f, ensure_ascii=False, indent=2)

        f.flush()

        os.fsync(f.fileno())

    os.replace(temp_file, state_file)


# ============================================================
# MENENTUKAN FILE OUTPUT
# ============================================================


def siapkan_output():
    """
    Menentukan file CSV dan state.

    Jika terdapat state sebelumnya,
    scraping akan dilanjutkan.

    Jika tidak ada,
    membuat file baru.
    """

    state_files = sorted(
        OUTPUT_DIR.glob("*.state.json"), key=lambda x: x.stat().st_mtime, reverse=True
    )

    # --------------------------------------------------------
    # CARI STATE YANG BELUM SELESAI
    # --------------------------------------------------------

    for state_file in state_files:

        state = baca_state(state_file)

        if not state:
            continue

        nama_file = Path(state["nama_file"])

        if not nama_file.is_absolute():

            nama_file = OUTPUT_DIR / nama_file

        # Jika state masih memiliki next_offset,
        # berarti scraping sebelumnya belum selesai.
        if state.get("next_offset", 0) >= 0:

            print("\n" + "=" * 70)
            print("RESUME SCRAPING")
            print("=" * 70)

            print("File CSV :", nama_file)

            print("Offset   :", f'{state.get("next_offset", 0):,}')

            print("Total    :", f'{state.get("total_data", 0):,}')

            return (nama_file, state_file, state)

    # --------------------------------------------------------
    # BUAT SCRAPING BARU
    # --------------------------------------------------------

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    nama_file = OUTPUT_DIR / f"hasil_scraping_{timestamp}.csv"

    state_file = OUTPUT_DIR / f"hasil_scraping_{timestamp}.state.json"

    state = buat_state_baru(nama_file)

    simpan_state(state_file, state)

    print("\n" + "=" * 70)
    print("SCRAPING BARU")
    print("=" * 70)

    print("File CSV :", nama_file)

    print("State    :", state_file)

    return (nama_file, state_file, state)


# ============================================================
# MEMULIHKAN FILE CSV
# ============================================================


def pulihkan_file_csv(nama_file, state):
    """
    Memulihkan ukuran CSV sesuai state terakhir.

    Ini penting untuk menghindari duplicate batch
    jika program mati setelah CSV selesai ditulis
    tetapi state belum sempat diperbarui.

    Jika file lebih besar dari ukuran yang tercatat
    di state, file akan dipotong kembali.
    """

    file_size = state.get("file_size", 0)

    if not nama_file.exists():

        return

    ukuran_sekarang = nama_file.stat().st_size

    if ukuran_sekarang > file_size:

        print("\n[RESUME] " "Memulihkan ukuran CSV...")

        print("Ukuran sekarang :", f"{ukuran_sekarang:,}", "bytes")

        print("Ukuran state   :", f"{file_size:,}", "bytes")

        with open(nama_file, "r+b") as f:

            f.truncate(file_size)

        print("[OK] CSV berhasil dipulihkan.")


# ============================================================
# MENULIS BATCH KE CSV
# ============================================================


def simpan_batch_csv(rows, nama_file, header_sudah_ditulis):
    """
    Menulis satu batch row langsung ke CSV.
    Tidak menggunakan DataFrame.
    Return:
        header_sudah_ditulis
        file_size
    """
    if not rows:
        return (
            header_sudah_ditulis,
            (nama_file.stat().st_size if nama_file.exists() else 0),
        )

    # --------------------------------------------------------
    # TENTUKAN HEADER
    # --------------------------------------------------------
    if header_sudah_ditulis:
        mode = "a"
        write_header = False
    else:
        mode = "w"
        write_header = True

    # --------------------------------------------------------
    # AMBIL KOLOM
    # --------------------------------------------------------

    # Data response Superset biasanya berupa dict.
    #
    # Ambil union semua key untuk mengantisipasi
    # adanya kolom tambahan.

    fieldnames = []
    field_set = set()
    for row in rows:
        if not isinstance(row, dict):
            continue
        for key in row.keys():
            if key not in field_set:
                field_set.add(key)
                fieldnames.append(key)
    if not fieldnames:
        print("[WARNING] " "Batch tidak memiliki data dictionary.")
        return (
            header_sudah_ditulis,
            (nama_file.stat().st_size if nama_file.exists() else 0),
        )

    # --------------------------------------------------------
    # TULIS CSV
    # --------------------------------------------------------

    with open(nama_file, mode, newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(
            f, fieldnames=fieldnames, extrasaction="ignore", restval=""
        )
        if write_header:
            writer.writeheader()
        for row in rows:
            if not isinstance(row, dict):
                continue
            writer.writerow(row)
        # Pastikan seluruh data benar-benar
        # masuk ke disk sebelum state disimpan.
        f.flush()
        os.fsync(f.fileno())
    return (True, nama_file.stat().st_size)


# ============================================================
# SCRAP SEMUA DATA
# ============================================================
def scrap_semua_data(context, csrf_token, nama_file, state_file, state):
    """
    Scraping streaming ke CSV.
    Tidak menyimpan semua data di RAM.
    Resume menggunakan OFFSET yang tersimpan
    pada state file.
    """
    # --------------------------------------------------------
    # PULIHKAN FILE
    # --------------------------------------------------------
    pulihkan_file_csv(nama_file, state)
    offset = state.get("next_offset", 0)
    total_data = state.get("total_data", 0)
    batch_ke = state.get("batch_ke", 0)
    header_sudah_ditulis = state.get("header_sudah_ditulis", False)
    print("\n" + "=" * 70)
    print("MEMULAI SCRAPING STREAMING")
    print("=" * 70)
    print(f"Mulai OFFSET : {offset:,}")
    print(f"Total awal   : {total_data:,}")
    print(f"File CSV     : {nama_file}")
    print("=" * 70)
    # ========================================================
    # LOOP OFFSET
    # ========================================================
    while True:
        print("\n")
        print("=" * 70)
        print("SCRAP DATA")
        print(f"Batch  : {batch_ke + 1:,}")
        print(f"LIMIT  : {QUERY_LIMIT:,}")
        print(f"OFFSET : {offset:,}")
        print("=" * 70)
        # ----------------------------------------------------
        # BUAT SQL
        # ----------------------------------------------------
        sql = buat_sql(offset)
        # ----------------------------------------------------
        # EXECUTE
        # ----------------------------------------------------
        result = execute_query(
            context=context, sql=sql, csrf_token=csrf_token, offset=offset
        )
        # ----------------------------------------------------
        # CSRF INVALID
        # ----------------------------------------------------
        if result == "REFRESH_CSRF":
            print("\n[ERROR] " "CSRF token sudah tidak valid.")
            print("Scraping dihentikan.")
            print("State sudah tersimpan sampai:")
            print(f"OFFSET {offset:,}")
            return (nama_file, total_data, False)
        # ----------------------------------------------------
        # REQUEST GAGAL
        # ----------------------------------------------------

        if result is None:
            print(f"\n[ERROR] " f"OFFSET {offset:,} gagal.")
            print("Scraping dihentikan.")
            print("Jalankan script kembali " "untuk melanjutkan.")
            return (nama_file, total_data, False)
        # ----------------------------------------------------
        # AMBIL ROW
        # ----------------------------------------------------
        rows = ambil_rows(result)
        jumlah = len(rows)
        print(f"\nJumlah row diterima: " f"{jumlah:,}")
        # ----------------------------------------------------
        # TIDAK ADA DATA
        # ----------------------------------------------------
        if jumlah == 0:
            print("\nTidak ada data lagi.")
            print("Scraping selesai.")
            # Tandai selesai dengan next_offset = -1
            state["next_offset"] = -1
            state["total_data"] = total_data
            state["batch_ke"] = batch_ke
            state["header_sudah_ditulis"] = header_sudah_ditulis
            state["file_size"] = nama_file.stat().st_size if nama_file.exists() else 0
            simpan_state(state_file, state)
            return (nama_file, total_data, True)

        # ----------------------------------------------------
        # SIMPAN BATCH
        # ----------------------------------------------------
        header_sudah_ditulis, file_size = simpan_batch_csv(
            rows=rows, nama_file=nama_file, header_sudah_ditulis=(header_sudah_ditulis)
        )
        # ----------------------------------------------------
        # UPDATE COUNTER
        # ----------------------------------------------------
        batch_ke += 1
        total_data += jumlah
        print(f"Batch ke       : {batch_ke:,}")
        print(f"Data batch     : {jumlah:,}")
        print(f"Total data     : {total_data:,}")
        print(f"Ukuran CSV     : " f"{file_size / (1024 * 1024):,.2f} MB")
        # ----------------------------------------------------
        # HALAMAN TERAKHIR
        # ----------------------------------------------------

        if jumlah < QUERY_LIMIT:
            print("\nJumlah row kurang dari " f"{QUERY_LIMIT:,}.")
            print("Ini adalah halaman terakhir.")
            # ------------------------------------------------
            # SIMPAN STATE SELESAI
            # ------------------------------------------------
            state["next_offset"] = -1
            state["total_data"] = total_data
            state["batch_ke"] = batch_ke
            state["header_sudah_ditulis"] = header_sudah_ditulis
            state["file_size"] = file_size
            simpan_state(state_file, state)
            print("\n[OK] State akhir tersimpan.")
            return (nama_file, total_data, True)
        # ----------------------------------------------------
        # NEXT OFFSET
        # ----------------------------------------------------
        next_offset = offset + QUERY_LIMIT
        # ----------------------------------------------------
        # SIMPAN STATE SETELAH BATCH BENAR-BENAR
        # MASUK KE DISK
        # ----------------------------------------------------
        state["next_offset"] = next_offset

        state["total_data"] = total_data

        state["batch_ke"] = batch_ke

        state["header_sudah_ditulis"] = header_sudah_ditulis

        state["file_size"] = file_size

        simpan_state(state_file, state)

        print("\n[STATE] " "Progress berhasil disimpan.")

        print(f"Next OFFSET : " f"{next_offset:,}")

        # ----------------------------------------------------
        # NEXT LOOP
        # ----------------------------------------------------

        offset = next_offset

        # Delay kecil agar request tidak terlalu rapat

        waktu_tunggu = random.uniform(1, 3)

        print(f"Menunggu " f"{waktu_tunggu:.1f} detik...")

        time.sleep(waktu_tunggu)


# ============================================================
# MAIN SCRAPING
# ============================================================


def jalankan_scraping():
    nama_file, state_file, state = siapkan_output()
    # --------------------------------------------------------
    # JIKA SUDAH SELESAI
    # --------------------------------------------------------
    if state.get("next_offset") == -1:
        print("\nScraping sebelumnya " "sudah selesai.")
        print("File:", nama_file)
        print("Total:", f'{state.get("total_data", 0):,}')
        return (nama_file, state.get("total_data", 0))
    # ========================================================
    # PLAYWRIGHT
    # ========================================================
    with sync_playwright() as p:
        # ----------------------------------------------------
        # BROWSER
        # ----------------------------------------------------
        browser = p.chromium.launch(
            headless=False, args=["--disable-blink-features=AutomationControlled"]
        )
        # ----------------------------------------------------
        # CONTEXT
        # ----------------------------------------------------
        context = browser.new_context(
            accept_downloads=True,
            user_agent=(
                "Mozilla/5.0 "
                "(Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 "
                "(KHTML, like Gecko) "
                "Chrome/126.0.0.0 "
                "Safari/537.36"
            ),
            viewport={"width": 1366, "height": 768},
            locale="id-ID",
            timezone_id="Asia/Jakarta",
        )

        page = context.new_page()

        Stealth().apply_stealth_sync(page)

        try:

            # =================================================
            # LOGIN
            # =================================================

            login_dashboard(page)

            # =================================================
            # PASANG PENANGKAP REQUEST
            # =================================================

            request_info = pasang_penangkap_csrf(page)

            # =================================================
            # BUKA SQL LAB
            # =================================================

            print("\nMembuka SQL Lab...")

            page.goto(LINK_SQLLAB_URL, wait_until="domcontentloaded")

            page.wait_for_timeout(5000)

            print("SQL Lab URL:", page.url)

            # =================================================
            # TUNGGU CSRF
            # =================================================

            csrf_token = tunggu_csrf(page, request_info)

            # =================================================
            # SCRAP
            # =================================================

            print("\nMemulai scraping otomatis...")

            nama_file, total_data, selesai = scrap_semua_data(
                context=context,
                csrf_token=csrf_token,
                nama_file=nama_file,
                state_file=state_file,
                state=state,
            )

            # =================================================
            # HASIL
            # =================================================

            print("\n" + "=" * 70)

            if selesai:

                print("SCRAPING SELESAI")

            else:

                print("SCRAPING BERHENTI / BELUM SELESAI")

            print("=" * 70)

            print(f"Total data : {total_data:,}")

            print(f"File CSV   : " f"{nama_file.resolve()}")

            print(f"State      : " f"{state_file.resolve()}")

            print("=" * 70)

            return (nama_file, total_data)

        finally:

            print("\nMenutup browser...")

            context.close()

            browser.close()


# ============================================================
# MAIN
# ============================================================


def main():

    print("\n")

    print("=" * 70)

    print("BOT SCRAPER SQL LAB FASIH")

    print("CSV STREAMING + RESUME")

    print("=" * 70)

    nama_file, total_data = jalankan_scraping()

    print("\n" + "=" * 70)

    print("SELESAI")

    print("=" * 70)

    print(f"Total data : {total_data:,}")

    print(f"File CSV   : " f"{Path(nama_file).resolve()}")

    print("=" * 70)


# ============================================================
# RUN
# ============================================================
if __name__ == "__main__":
    main()
