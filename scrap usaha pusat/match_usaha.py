#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
match_usaha.py  (versi fuzzy matching)

Mencocokkan usaha kandidat scraping (tutup, tidak ditemukan, duplikat) dengan usaha
berkode lain di list yang sama, lalu menentukan target scraping.

Satu file dapat memuat list UB dan non UB sekaligus; keduanya diproses dalam
satu kali jalan dengan OUTPUT TERPISAH. Tiap kandidat dikenali dari flag-nya:
non UB (default kode 3,0) dan UB (default kode 5,6,7) -- kode kandidat kedua
frame tidak beririsan. Pool pembanding = usaha AKTIF (flag di luar semua kode
kandidat) dan dipakai BERSAMA oleh kedua frame.

Alur (dijalankan per frame: non UB lalu UB):
  1. Filter usaha kandidat scraping: kode_flag_keberadaan dalam kode frame tsb:
     - non UB: kode 3 dan 0 (tidak ditemukan / tutup)
     - UB    : kode 5, 6, dan 7 (tutup / tidak ditemukan / duplikat)
  2. Pool pembanding = usaha aktif (flag BUKAN kode kandidat non UB maupun UB),
     dipakai bersama kedua frame.
  3. Kandidat hanya dicari di KECAMATAN YANG SAMA
     (kd_prov + kd_kab + kd_kec sama). Kecamatan berbeda = tidak match.
  4. Di dalam kecamatan, nama dan alamat dibandingkan secara FUZZY
     (toleran typo, urutan kata, singkatan alamat, nomor rumah):
       - skor_nama   >= ambang nama   (default 88)
       - skor_alamat >= ambang alamat (default 80)
     Lalu:
       Rule 1: kode wilayah sama (kd_prov+kd_kab+kd_kec+kd_desa)
               -> status_match = "match"
       Rule 2: kode wilayah beda (desa lain, kecamatan sama)
               -> status_match = "match di wilayah lain"
       Tidak ada yang cocok -> "tidak match"
     Kolom idsbr_usaha_match berisi idsbr usaha yang cocok (diurutkan dari
     skor tertinggi, pisah ";" bila banyak); usaha_match_kode_flag_keberadaan dan
     usaha_match_wilayah berisi kode_flag_keberadaan dan kode wilayah usaha tsb
     (urutan sama).
  5. Usaha kandidat scraping yang "tidak match" = TARGET SCRAPING.
  6. SELEKSI USAHA vs PEMILIK (khusus frame NON UB):
     Nama diklasifikasikan dengan SKOR DUA SISI, bukan first-match-wins.
     skor_usaha dan skor_pemilik dihitung terpisah lalu dibandingkan:
       - sinyal usaha : badan usaha di awal (PT/CV/UD...), kata kunci usaha
                        (toko/warung/bengkel...), >=2 kata "kemakmuran"
                        berdampingan (Sumber Rezeki, Barokah Jaya), angka/&,
                        nama sangat panjang.
       - sinyal pemilik: token yang ada di gazetteer nama depan
                        (file nama_depan.txt; lihat --file-nama).
     Hasil: "usaha" / "pemilik" / "ragu". "ragu" = kedua sisi lemah; nama tak
     bisa dipastikan. Ini mencegah usaha tanpa kata kunci ("Barokah Jaya")
     divonis "pemilik" lalu hilang dari target secara diam-diam.
     Output:
       - kolom klasifikasi_nama (+ skor_usaha, skor_pemilik untuk audit)
         pada hasil_match.csv / sheet Semua
       - file scraping CSV *_target_scraping_usaha.csv berisi usaha + ragu
         (lebih baik scrape kelebihan daripada melewatkan usaha); "pemilik"
         dikecualikan. Sheet Target_Usaha / Target_Ragu / Target_Pemilik
         memisahkan ketiganya di file Excel untuk pengecekan.
     CATATAN: sinyal pemilik hanya dari gazetteer; bila file nama tidak ada,
     hasil hanya "usaha"/"ragu". Untuk UB seleksi tidak dijalankan
     (bisa dipaksa dengan --seleksi-pemilik ya).

Dependensi:  
    pip install rapidfuzz

Pemakaian:
    python match_usaha.py input.csv hasil_match
    python match_usaha.py input.csv hasil_match --ambang-nama 92 --ambang-alamat 85
    python match_usaha.py input.csv hasil_match --kode-ub ""     # hanya proses non UB
    python match_usaha.py input.csv hasil_match --kode-nonub ""  # hanya proses UB

Output (dengan prefix "hasil_match") -- kolom 'frame' penanda UB/non UB:
    hasil_match.csv                        -> semua kandidat (UB + non UB) + hasil
    hasil_match_target_scraping_usaha.csv  -> target scraping: non UB usaha + UB (file utama)
    hasil_match_target_scraping_ragu.csv   -> target scraping: non UB nama ambigu (cek dulu)
    hasil_match.xlsx                       -> 4 sheet: Semua, Target_Usaha,
                                              Target_Ragu, Pemilik_Dikecualikan
Status match / match wilayah lain dibaca dari kolom status_match (+ filter),
tidak dipisah jadi sheet/ file sendiri.
"""

import argparse
import csv
import re
import sys
import unicodedata
from collections import defaultdict
from difflib import SequenceMatcher

try:
    from rapidfuzz import fuzz as _rf_fuzz
except ImportError:  
    _rf_fuzz = None

# ============================================================
# CONFIG
# ============================================================

# Kode kandidat scraping per frame. UB & non UB kini diproses sekaligus dari
# SATU file; tiap kandidat dikenali dari flag-nya. Set "" untuk melewati frame.
KODE_NONUB_DEFAULT = "3,0"      # non UB: tidak ditemukan / tutup
KODE_UB_DEFAULT = "5,6,7"       # UB: tutup / tidak ditemukan / duplikat
AMBANG_NAMA_DEFAULT = 88.0
AMBANG_ALAMAT_DEFAULT = 80.0

# File gazetteer nama depan (dibuat oleh build_gazetteer.py). Default: cari
# di samping script. Bisa diganti lewat --file-nama.
FILE_NAMA_DEFAULT = "nama_depan.txt"

STATUS_MATCH = "match"
STATUS_WILAYAH_LAIN = "match di wilayah lain"
STATUS_TIDAK = "tidak match"

# Alias nama kolom (dipilih yang pertama ditemukan).
COL_ALIASES = {
    "idsbr": ["idsbr", "id_sbr", "id"],
    "nama": ["nama", "nama_perusahaan", "nama_se", "nama_bpjph", "nama_usaha"],
    "alamat": ["alamat", "address", "alamat_se", "alamat_bpjph", "alamat_usaha"],
    "kd_prov": ["kd_prov", "kode_prov", "kdprov"],
    "kd_kab": ["kd_kab", "kode_kab", "kdkab"],
    "kd_kec": ["kd_kec", "kode_kec", "kdkec"],
    "kd_desa": ["kd_desa", "kode_desa", "kddesa"],
    "flag": ["kode_flag_keberadaan", "keberadaan"],
}

# Kode kode_flag_keberadaan khusus UB: bila ada di kode kandidat -> mode UB.
KODE_UB = {"5", "6", "7"}

# Kata kunci umum yang menandakan sebuah entitas usaha
# (disalin dari klasifikasi_usaha_pemilik.py agar script ini tidak butuh pandas).
KATA_KUNCI_USAHA = [
    'pt', 'cv', 'ud', 'toko', 'warung', 'kios', 'restoran', 'resto',
    'cafe', 'kafe', 'bengkel', 'salon', 'laundry', 'apotek', 'klinik',
    'butik', 'grup', 'group', 'perusahaan', 'firma', 'koperasi',
    'yayasan', 'bakery', 'catering', 'studio', 'agency', 'agen',
    'distributor', 'supplier', 'mart', 'swalayan', 'minimarket',
    'showroom', 'gallery', 'galeri', 'jasa', 'servis', 'service'
]

ADDED_COLUMNS = [
    "kode_wilayah",
    "status_match",          # <- sheet/CSV target_scraping berhenti sampai kolom ini
    "idsbr_usaha_match",
    "usaha_match_kode_flag_keberadaan",
    "usaha_match_wilayah",
    "skor_nama",
    "skor_alamat",
]
LAST_TARGET_COLUMN = "status_match"
KLASIFIKASI_COLUMN = "klasifikasi_nama"   # hanya untuk non UB (ditaruh paling akhir)
SKOR_USAHA_COLUMN = "skor_usaha"          # transparansi keputusan klasifikasi
SKOR_PEMILIK_COLUMN = "skor_pemilik"

# --- Klasifikasi usaha vs pemilik: sinyal skoring ---
KATA_KUNCI_USAHA_SET = set(KATA_KUNCI_USAHA)

# Badan usaha di AWAL nama = sinyal usaha terkuat.
BADAN_AWAL = {"pt", "cv", "ud", "pd", "tb", "fa", "koperasi", "kop"}

# Akhiran lengket yang menandakan ritel (alfamart, indomaret-> "mart"/"maret").
STICKY_USAHA = ("mart", "maret", "swalayan", "grosir")

# Kata "kemakmuran": khas penamaan usaha tanpa kata kunci badan usaha.
# SATU kata = lemah (bisa nama orang); DUA+ kata berdampingan = kuat.
KATA_USAHA_ABSTRAK = {
    "jaya", "makmur", "sejahtera", "berkah", "barokah", "rezeki", "rejeki",
    "sumber", "abadi", "mandiri", "sentosa", "mulia", "agung", "karya",
    "utama", "bersama", "indah", "sukses", "maju", "lancar", "berkat",
    "amanah", "sejati", "perkasa", "gemilang", "cemerlang", "subur",
    "makmur", "berkah", "rahayu", "santosa", "waras", "lestari", "jaya",
    "harapan", "sinar", "cahaya", "bintang", "surya", "mentari", "putra",
    "putri", "family", "famili", "saudara", "sahabat", "mitra", "anugerah",
    "anugrah", "rahmat", "rizki", "rejeki", "berdikari", "prima", "unggul",
}

# Gazetteer nama depan diisi dari file saat runtime (lihat muat_gazetteer).
NAMA_DEPAN = set()

# True: 1 nama gazetteer + tanpa sinyal usaha -> pemilik (tangkap nama orang
# ber-nama-belakang tak lazim, mengecilkan file scraping). False: konservatif,
# butuh >=2 sinyal pemilik; nama begitu jatuh ke 'ragu' dan tetap discrape.
AMBANG_PEMILIK_LONGGAR = True

# Badan usaha yang diabaikan saat membandingkan nama.
LEGAL_FORMS = {"pt", "cv", "ud", "pd", "tb", "fa", "firma", "persero", "tbk"}

# Singkatan alamat -> bentuk baku.
ADDRESS_ABBR = {
    "jl": "jalan", "jln": "jalan",
    "dsn": "dusun", "dus": "dusun",
    "ds": "desa", "kel": "kelurahan", "kec": "kecamatan",
    "kp": "kampung", "kmp": "kampung",
    "gg": "gang", "gp": "gang",
    "kav": "kavling", "kompl": "komplek", "perum": "perumahan",
    "no": "nomor", "nmr": "nomor",
}

# Kata alamat yang terlalu umum (tidak membedakan) -> dibuang dari skor.
ADDRESS_NOISE = {
    "jalan", "nomor", "desa", "kelurahan", "kecamatan", "kabupaten", "kota",
    "dusun", "kampung", "indonesia", "provinsi",
}


# ============================================================
# NORMALISASI
# ============================================================

def _ascii_lower(s):
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return s.lower().replace("&", " dan ")


def norm_text(s):
    """Lowercase, buang aksen & tanda baca, rapikan spasi."""
    s = re.sub(r"[^a-z0-9]+", " ", _ascii_lower(s))
    return re.sub(r"\s+", " ", s).strip()


def norm_name(s, abaikan_badan_usaha=True):
    """Normalisasi nama: token terurut, tanpa badan usaha."""
    toks = norm_text(s).split()
    if abaikan_badan_usaha:
        toks = [t for t in toks if t not in LEGAL_FORMS]
    return " ".join(toks)


def _strip_zeros(tok):
    return tok.lstrip("0") or "0" if tok.isdigit() else tok


def norm_address(s):
    """Normalisasi alamat: singkatan dibakukan, RT/RW dibuang, noise dibuang.

    Mengembalikan (teks_ternormalisasi, himpunan_nomor).
    """
    t = _ascii_lower(s)
    t = re.sub(r"[^a-z0-9]+", " ", t)
    t = re.sub(r"(?<=[a-z])(?=\d)", " ", t)          # no5 -> no 5, rt003 -> rt 003
    t = re.sub(r"\b(rt|rw)\s*\d+\b", " ", t)         # buang "rt 003", "rw 02"
    t = re.sub(r"\s+", " ", t).strip()

    toks = []
    for tok in t.split():
        tok = ADDRESS_ABBR.get(tok, tok)
        tok = _strip_zeros(tok)
        if tok in ADDRESS_NOISE:
            continue
        toks.append(tok)

    numbers = {tok for tok in toks if tok.isdigit()}
    return " ".join(toks), numbers


def muat_gazetteer(path):
    """Isi NAMA_DEPAN dari file (satu nama per baris; '#' & kosong diabaikan).

    Mengembalikan jumlah nama yang dimuat. Bila file tidak ada, NAMA_DEPAN
    tetap kosong dan klasifikasi otomatis jatuh ke mode usaha/ragu
    (sinyal pemilik mati) -- lihat catatan di klasifikasi_nama.
    """
    import os
    NAMA_DEPAN.clear()
    if not path or not os.path.exists(path):
        return 0
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            tok = line.strip()
            if tok and not tok.startswith("#"):
                NAMA_DEPAN.add(norm_text(tok))
    NAMA_DEPAN.discard("")
    return len(NAMA_DEPAN)


def klasifikasi_nama(nama):
    """Klasifikasi 'usaha' / 'pemilik' / 'ragu' dengan skor dua sisi.

    Alih-alih memvonis 'pemilik' hanya karena tak ada kata kunci usaha
    (rawan: 'Barokah Jaya' bukan nama orang), skor_usaha dan skor_pemilik
    dihitung terpisah lalu dibandingkan. Bila sama-sama lemah -> 'ragu',
    supaya tidak ada target scraping yang hilang diam-diam.

    Catatan: sinyal 'pemilik' bersumber HANYA dari gazetteer NAMA_DEPAN
    (data ini tanpa gelar/patronim). Bila gazetteer kosong, skor_pemilik
    selalu 0 -> hasil hanya 'usaha' atau 'ragu'.

    Mengembalikan (label, skor_usaha, skor_pemilik).
    """
    toks = norm_text(nama).split()
    if not toks:
        return "ragu", 0, 0

    skor_usaha = skor_pemilik = 0

    # Kata kunci usaha = penentu mutlak. Bila ada (toko/warung/bengkel/PT/CV...),
    # langsung 'usaha' tanpa adu skor -- ini sinyal paling tegas.
    if toks[0] in BADAN_AWAL or any(t in KATA_KUNCI_USAHA_SET for t in toks):
        return "usaha", 99, 0

    # --- Sinyal USAHA (untuk nama tanpa kata kunci eksplisit) ---
    if any(t.endswith(STICKY_USAHA) and len(t) > 4 for t in toks):
        skor_usaha += 2
    n_abstrak = sum(1 for t in toks if t in KATA_USAHA_ABSTRAK)
    if n_abstrak >= 2:
        skor_usaha += 2          # dua kata abstrak berdampingan -> usaha
    elif n_abstrak == 1:
        skor_usaha += 1          # satu kata -> lemah (bisa jadi nama orang)
    if any(c.isdigit() for c in (nama or "")) or "&" in (nama or ""):
        skor_usaha += 1
    if len(toks) >= 5:
        skor_usaha += 1          # nama sangat panjang cenderung usaha

    # --- Sinyal PEMILIK (hanya dari gazetteer nama) ---
    n_nama = sum(1 for t in toks if t in NAMA_DEPAN)
    skor_pemilik += min(n_nama, 2)          # +1 per token nama, maks +2

    # --- Keputusan: default 'ragu' ---
    if skor_usaha >= 2 and skor_usaha > skor_pemilik:
        return "usaha", skor_usaha, skor_pemilik
    if skor_pemilik >= 2 and skor_pemilik > skor_usaha:
        return "pemilik", skor_usaha, skor_pemilik
    # Longgar (opsional): 1 token nama gazetteer + TANPA sinyal usaha sama
    # sekali -> pemilik. Menangkap "Siti Nurhaliza" (nama depan umum + nama
    # belakang tak lazim) tanpa menyeret nama usaha, yang hampir selalu
    # memicu minimal satu sinyal usaha. Dimatikan dengan AMBANG_PEMILIK_LONGGAR
    # = False di CONFIG bila ingin konservatif.
    if AMBANG_PEMILIK_LONGGAR and skor_pemilik >= 1 and skor_usaha == 0:
        return "pemilik", skor_usaha, skor_pemilik
    return "ragu", skor_usaha, skor_pemilik


def norm_code(s):
    """Samakan kode flag: '3', '3.0' -> '3'; kosong tetap kosong."""
    v = str(s).strip() if s is not None else ""
    if not v:
        return ""
    try:
        return str(int(float(v)))
    except ValueError:
        return v.lower()


def pick(row, key):
    for col in COL_ALIASES[key]:
        if col in row and row[col] is not None:
            return str(row[col]).strip()
    return ""


# ============================================================
# SKOR KEMIRIPAN (0 - 100)
# ============================================================

def _sort_tokens(s):
    return " ".join(sorted(s.split()))


def _ratio(a, b):
    if not a or not b:
        return 0.0
    if _rf_fuzz is not None:
        return float(_rf_fuzz.ratio(a, b))
    return 100.0 * SequenceMatcher(None, a, b).ratio()


def score_name(a, b):
    """Kemiripan nama (urutan kata tidak berpengaruh, toleran typo)."""
    if not a or not b:
        return 0.0
    return _ratio(_sort_tokens(a), _sort_tokens(b))


def score_address(a, b):
    """Kemiripan alamat. a/b = (teks, himpunan_nomor) dari norm_address."""
    ta, na = a
    tb, nb = b
    if not ta or not tb:
        return 0.0

    score = _ratio(_sort_tokens(ta), _sort_tokens(tb))

    # Satu alamat lebih lengkap dari yang lain: pakai containment (maks 95).
    A, B = set(ta.split()), set(tb.split())
    if min(len(A), len(B)) >= 2:
        containment = 100.0 * len(A & B) / min(len(A), len(B))
        score = max(score, min(containment, 95.0))

    # Nomor rumah/blok: bila kedua sisi punya nomor tetapi tidak ada yang sama,
    # kemungkinan lokasi berbeda -> skor dipotong.
    if na and nb and not (na & nb):
        score *= 0.6

    return score


# ============================================================
# I/O
# ============================================================

def read_csv(path):
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        sample = f.read(4096)
        f.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        except csv.Error:
            dialect = csv.excel
        reader = csv.DictReader(f, dialect=dialect)
        rows = list(reader)
        return rows, list(reader.fieldnames or [])


def write_csv(path, rows, fieldnames):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)


def write_xlsx(path, sheets):
    """sheets = [(judul, rows, fieldnames), ...]"""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font
    except ImportError:
        print("openpyxl tidak terpasang, file Excel dilewati "
              "(pip install openpyxl).")
        return False

    wb = Workbook()
    for i, (title, rows, fieldnames) in enumerate(sheets):
        ws = wb.active if i == 0 else wb.create_sheet()
        ws.title = title
        ws.append(fieldnames)
        for c in ws[1]:
            c.font = Font(bold=True)
        for r in rows:
            ws.append([r.get(c, "") for c in fieldnames])
        ws.freeze_panes = "A2"
        if ws.dimensions:
            ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            width = min(max(len(str(c.value or "")) for c in col) + 2, 50)
            ws.column_dimensions[col[0].column_letter].width = width
    wb.save(path)
    return True


# ============================================================
# MATCHING
# ============================================================

def kode_wilayah(row):
    """Kode wilayah = kd_prov + kd_kab + kd_kec + kd_desa (digabung apa adanya)."""
    return "".join(pick(row, k) for k in ("kd_prov", "kd_kab", "kd_kec", "kd_desa"))


def kode_kecamatan(row):
    """Kode kecamatan = kd_prov + kd_kab + kd_kec."""
    return "".join(pick(row, k) for k in ("kd_prov", "kd_kab", "kd_kec"))


def _name_keys(name_norm):
    """Kunci blocking: 4 huruf pertama tiap token (toleran typo di akhir kata)."""
    return {t[:4] for t in name_norm.split() if len(t) >= 2}


def prepare(row, abaikan_badan_usaha):
    """Hitung fitur ternormalisasi sekali per baris."""
    name = norm_name(pick(row, "nama"), abaikan_badan_usaha)
    addr = norm_address(pick(row, "alamat"))
    return {
        "row": row,
        "id": pick(row, "idsbr"),
        "name": name,
        "addr": addr,
        "wil": kode_wilayah(row),
        "kec": kode_kecamatan(row),
        "keys": _name_keys(name),
    }


def cocokkan(kandidat, index, seleksi, args):
    """Cocokkan daftar kandidat satu frame terhadap index pool. Return list hasil."""
    hasil = []
    for b in kandidat:
        out = dict(b["row"])
        out["kode_wilayah"] = b["wil"]

        hits = []  # (rule, skor_gabungan, skor_nama, skor_alamat, pool_item)

        if b["name"] and b["addr"][0] and b["kec"]:
            seen = set()
            for k in b["keys"]:
                for c in index.get(b["kec"], {}).get(k, []):
                    if id(c) in seen or (c["id"] and c["id"] == b["id"]):
                        continue
                    seen.add(id(c))

                    sn = score_name(b["name"], c["name"])
                    if sn < args.ambang_nama:
                        continue
                    sa = score_address(b["addr"], c["addr"])
                    if sa < args.ambang_alamat:
                        continue

                    rule = 1 if c["wil"] == b["wil"] else 2
                    hits.append((rule, (sn + sa) / 2, sn, sa, c))

        rule1 = sorted((h for h in hits if h[0] == 1), key=lambda h: -h[1])
        rule2 = sorted((h for h in hits if h[0] == 2), key=lambda h: -h[1])

        if rule1:                                  # Rule 1
            out["status_match"] = STATUS_MATCH
            chosen = rule1
        elif rule2:                                # Rule 2
            out["status_match"] = STATUS_WILAYAH_LAIN
            chosen = rule2
        else:
            out["status_match"] = STATUS_TIDAK
            chosen = []

        out["idsbr_usaha_match"] = ";".join(h[4]["id"] for h in chosen)
        out["usaha_match_kode_flag_keberadaan"] = ";".join(
            pick(h[4]["row"], "flag") for h in chosen
        )
        out["usaha_match_wilayah"] = ";".join(h[4]["wil"] for h in chosen)
        out["skor_nama"] = round(chosen[0][2], 1) if chosen else ""
        out["skor_alamat"] = round(chosen[0][3], 1) if chosen else ""
        if seleksi:
            label, su, sp = klasifikasi_nama(pick(b["row"], "nama"))
            out[KLASIFIKASI_COLUMN] = label
            out[SKOR_USAHA_COLUMN] = su
            out[SKOR_PEMILIK_COLUMN] = sp
        hasil.append(out)
    return hasil


def ikut_scrape(r):
    """True bila baris masuk file target scraping.

    - UB: semua target (status 'tidak match') ikut.
    - non UB: hanya 'usaha'/'ragu' yang ikut; 'pemilik' dikecualikan. Bila
      seleksi dimatikan (tak ada klasifikasi), semua target non UB ikut.
    """
    if r["status_match"] != STATUS_TIDAK:
        return False
    if r.get("frame") == "UB":
        return True
    klas = r.get(KLASIFIKASI_COLUMN)
    if not klas:                       # seleksi mati / tak terklasifikasi
        return True
    return klas in ("usaha", "ragu")


def run(args):
    rows, fieldnames = read_csv(args.input_csv)
    if not rows:
        print("CSV kosong.")
        return 1

    missing = [
        k for k in ("idsbr", "nama", "alamat", "flag")
        if not any(c in fieldnames for c in COL_ALIASES[k])
    ]
    if missing:
        print("Kolom wajib tidak ditemukan:", ", ".join(missing))
        print("Kolom yang ada:", ", ".join(fieldnames))
        return 1

    kode_nonub = {norm_code(x) for x in args.kode_nonub.split(",") if x.strip()}
    kode_ub = {norm_code(x) for x in args.kode_ub.split(",") if x.strip()}
    if not kode_nonub and not kode_ub:
        print("Tidak ada kode kandidat (non UB & UB kosong). Tidak ada yang "
              "diproses.")
        return 1
    tumpang = kode_nonub & kode_ub
    if tumpang:
        print(f"PERINGATAN: kode {sorted(tumpang)} ada di non UB DAN UB; baris "
              "dengan kode itu akan masuk ke non UB saja.")
        kode_ub = kode_ub - kode_nonub

    abaikan_bu = not args.pertahankan_badan_usaha

    # Seleksi usaha/pemilik: aktif untuk frame non UB (kecuali dipaksa tidak);
    # tidak aktif untuk UB (kecuali dipaksa ya).
    seleksi_nonub = args.seleksi_pemilik != "tidak"
    seleksi_ub = args.seleksi_pemilik == "ya"

    # Muat gazetteer nama depan bila ada frame yang pakai seleksi.
    import os
    if (kode_nonub and seleksi_nonub) or (kode_ub and seleksi_ub):
        fp = args.file_nama
        if not os.path.isabs(fp) and not os.path.exists(fp):
            cand = os.path.join(os.path.dirname(os.path.abspath(__file__)), fp)
            if os.path.exists(cand):
                fp = cand
        if muat_gazetteer(fp) == 0:
            print(f"PERINGATAN: gazetteer nama depan '{args.file_nama}' tidak "
                  "ditemukan/kosong. Sinyal 'pemilik' nonaktif -> hasil hanya "
                  "'usaha'/'ragu'. Jalankan build_gazetteer.py atau beri "
                  "--file-nama.\n")

    # --- Pisahkan kandidat per frame vs pool aktif (bersama) ---
    kand_nonub, kand_ub, pool = [], [], []
    for r in rows:
        kode = norm_code(pick(r, "flag"))
        if kode in kode_nonub:
            kand_nonub.append(prepare(r, abaikan_bu))
        elif kode in kode_ub:
            kand_ub.append(prepare(r, abaikan_bu))
        else:
            pool.append(prepare(r, abaikan_bu))     # usaha aktif = pool bersama

    # --- Blocking: kecamatan -> kunci nama -> daftar usaha pool (sekali) ---
    index = defaultdict(lambda: defaultdict(list))
    for p in pool:
        if p["name"] and p["addr"][0]:
            for k in p["keys"]:
                index[p["kec"]][k].append(p)

    # --- Matching kedua frame, lalu digabung (kolom 'frame' penanda asal) ---
    hasil_nonub = cocokkan(kand_nonub, index, seleksi_nonub, args)
    for r in hasil_nonub:
        r["frame"] = "non UB"
    hasil_ub = cocokkan(kand_ub, index, seleksi_ub, args)
    for r in hasil_ub:
        r["frame"] = "UB"
    hasil = hasil_nonub + hasil_ub

    # Susunan kolom: field asli + frame + kolom tambahan + kolom klasifikasi.
    added = ["frame"] + ADDED_COLUMNS + [
        KLASIFIKASI_COLUMN, SKOR_USAHA_COLUMN, SKOR_PEMILIK_COLUMN]
    out_fields = fieldnames + [c for c in added if c not in fieldnames]

    # File scraping (ringkas): identitas + frame + wilayah + status + klasifikasi.
    scrape_cols = ["frame", "kode_wilayah", "status_match", KLASIFIKASI_COLUMN]
    scrape_fields = fieldnames + [c for c in scrape_cols if c not in fieldnames]

    # Target scraping dipecah: 'ragu' (nama ambigu, cek dulu) vs 'usaha'
    # (termasuk UB & non UB tanpa klasifikasi). 'pemilik' dikecualikan.
    target_scrape = [r for r in hasil if ikut_scrape(r)]
    target_ragu = [r for r in target_scrape if r.get(KLASIFIKASI_COLUMN) == "ragu"]
    target_usaha = [r for r in target_scrape if r.get(KLASIFIKASI_COLUMN) != "ragu"]
    pemilik_excl = [
        r for r in hasil
        if r["status_match"] == STATUS_TIDAK
        and r.get(KLASIFIKASI_COLUMN) == "pemilik"
    ]

    prefix = args.output_prefix
    if prefix.lower().endswith(".csv"):
        prefix = prefix[:-4]
    path_hasil = prefix + ".csv"
    path_usaha = prefix + "_target_scraping_usaha.csv"
    path_ragu = prefix + "_target_scraping_ragu.csv"
    path_xlsx = prefix + ".xlsx"

    write_csv(path_hasil, hasil, out_fields)
    write_csv(path_usaha, target_usaha, scrape_fields)
    write_csv(path_ragu, target_ragu, scrape_fields)

    # 4 sheet, mudah dipahami:
    #   Semua                 -> semua kandidat + hasil (filter status_match / frame)
    #   Target_Usaha          -> target scraping: non UB usaha + UB (file utama)
    #   Target_Ragu           -> target scraping: non UB nama ambigu (cek dulu)
    #   Pemilik_Dikecualikan  -> bernama orang, tidak discrape (untuk cek)
    sheets = [
        ("Semua", hasil, out_fields),
        ("Target_Usaha", target_usaha, scrape_fields),
        ("Target_Ragu", target_ragu, scrape_fields),
        ("Pemilik_Dikecualikan", pemilik_excl, scrape_fields),
    ]
    xlsx_ok = write_xlsx(path_xlsx, sheets)

    # --- Ringkasan ---
    def hitung(frame):
        sub = [r for r in hasil if r["frame"] == frame]
        tgt = [r for r in sub if r["status_match"] == STATUS_TIDAK]
        return len(sub), len(tgt)

    print("=" * 60)
    print("MATCHING USAHA (fuzzy) -- UB + non UB, satu jalan")
    print("=" * 60)
    print(f"Total baris input            : {len(rows)}")
    print(f"Kode non UB / UB             : "
          f"{','.join(sorted(kode_nonub)) or '-'} / "
          f"{','.join(sorted(kode_ub)) or '-'}")
    print(f"Ambang skor nama / alamat    : {args.ambang_nama:g} / {args.ambang_alamat:g}")
    print(f"Mesin skor                   : {'rapidfuzz' if _rf_fuzz else 'difflib'}")
    print(f"Pool pembanding (aktif)      : {len(pool)}  (dipakai bersama)")
    print("-" * 60)
    if kode_nonub:
        n, t = hitung("non UB")
        print(f"non UB : {n} kandidat -> {t} tidak match")
    if kode_ub:
        n, t = hitung("UB")
        print(f"UB     : {n} kandidat -> {t} tidak match")
    print(f"Target scraping - usaha      : {len(target_usaha)}  (+ UB)")
    print(f"Target scraping - ragu       : {len(target_ragu)}  (cek dulu)")
    print(f"Pemilik dikecualikan         : {len(pemilik_excl)}")
    if NAMA_DEPAN:
        print(f"Gazetteer nama depan         : {len(NAMA_DEPAN)} nama")
    print("-" * 60)
    print("Output:", path_hasil, "(semua hasil)")
    print("Output:", path_usaha, "(scraping: non UB usaha + UB)")
    print("Output:", path_ragu, "(scraping: non UB ragu)")
    if xlsx_ok:
        print("Output:", path_xlsx, "(4 sheet)")
    print("=" * 60)
    return 0


def main():
    p = argparse.ArgumentParser(
        description="Matching fuzzy usaha kandidat (UB + non UB) dengan usaha "
                    "aktif; output terpisah per frame."
    )
    p.add_argument("input_csv", help="CSV input (daftar usaha, boleh campur UB & non UB)")
    p.add_argument("output_prefix",
                   help="Prefix file output, mis. 'hasil_match' (jadi "
                        "hasil_match.csv, *_target_scraping_usaha.csv, "
                        "*_target_scraping_ragu.csv, hasil_match.xlsx)")
    p.add_argument(
        "--kode-nonub",
        default=KODE_NONUB_DEFAULT,
        help="kode_flag_keberadaan kandidat non UB, pisah koma "
             f"(default '{KODE_NONUB_DEFAULT}'). Kosongkan (\"\") untuk lewati.",
    )
    p.add_argument(
        "--kode-ub",
        default=KODE_UB_DEFAULT,
        help="kode_flag_keberadaan kandidat UB, pisah koma "
             f"(default '{KODE_UB_DEFAULT}'). Kosongkan (\"\") untuk lewati.",
    )
    p.add_argument(
        "--ambang-nama",
        type=float,
        default=AMBANG_NAMA_DEFAULT,
        help=f"Skor minimum kemiripan nama 0-100 (default {AMBANG_NAMA_DEFAULT:g})",
    )
    p.add_argument(
        "--ambang-alamat",
        type=float,
        default=AMBANG_ALAMAT_DEFAULT,
        help=f"Skor minimum kemiripan alamat 0-100 (default {AMBANG_ALAMAT_DEFAULT:g})",
    )
    p.add_argument(
        "--pertahankan-badan-usaha",
        action="store_true",
        help="Jangan abaikan PT/CV/UD/PD/TB/FA saat membandingkan nama",
    )
    p.add_argument(
        "--seleksi-pemilik",
        choices=["auto", "ya", "tidak"],
        default="auto",
        help="Seleksi usaha vs pemilik. 'auto' (default): aktif untuk frame "
             "non UB, nonaktif untuk UB. 'ya': paksa aktif di kedua frame. "
             "'tidak': matikan di kedua frame.",
    )
    p.add_argument(
        "--file-nama",
        default=FILE_NAMA_DEFAULT,
        help="File gazetteer nama depan (satu nama per baris) untuk mengenali "
             f"nama pemilik (default '{FILE_NAMA_DEFAULT}' di samping script). "
             "Dibuat oleh build_gazetteer.py.",
    )
    return run(p.parse_args())


if __name__ == "__main__":
    sys.exit(main())
