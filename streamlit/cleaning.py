import io
import sqlite3
from contextlib import closing
from datetime import date, timedelta
from html import escape
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Monitoring Usaha", layout="wide")

# ============================================================
# Konfigurasi
# ============================================================
BASE_DIR = Path(__file__).resolve().parent
DB_PATH = (
    BASE_DIR
    / ".."
    / "cleansing_data"
    / "3 Oktober 2026"
    / "Mikro"
    / "RantabDBNAS"
    / "cleaning.db"
).resolve()

if not DB_PATH.exists():
    st.error(f"Database tidak ditemukan: {DB_PATH}")
    st.stop()

TABEL = "usaha"
SEMUA = "(Semua)"
KOLOM_FILTER = ("level_2_code", "kategori")  # whitelist kolom yang boleh difilter/group
LABEL = {"level_2_code": "kode wilayah", "kategori": "kategori"}

# ---------- Filter pengecekan ----------
AMBANG_OUTPUT = 1_000_000_000  # output >= 1 miliar
AMBANG_NTB = -5_000_000  # nilai tambah <= -5 juta

OPSI_SEMUA_DATA = "Semua data"
OPSI_CEK_1 = f"Output >= {AMBANG_OUTPUT:,} & flag_rasio_3 = 1"
OPSI_CEK_2 = f"Nilai tambah <= {AMBANG_NTB:,} & flag_rasio_3 = 1"
OPSI_CEK = [OPSI_SEMUA_DATA, OPSI_CEK_1, OPSI_CEK_2]

# CAST agar aman kalau flag tersimpan '1' / '1.0' (TEXT)
FLAG_3 = "CAST(flag_rasio_3_ntb_output AS REAL) = 1"
KONDISI_CEK = {
    OPSI_CEK_1: f"metrik_output >= {AMBANG_OUTPUT} AND {FLAG_3}",
    OPSI_CEK_2: f"metrik_nilai_tambah <= {AMBANG_NTB} AND {FLAG_3}",
}
DESKRIPSI_CEK = {
    OPSI_CEK_1: f"Hanya baris dengan metrik_output >= {AMBANG_OUTPUT:,} dan flag_rasio_3_ntb_output = 1.",
    OPSI_CEK_2: f"Hanya baris dengan metrik_nilai_tambah <= {AMBANG_NTB:,} dan flag_rasio_3_ntb_output = 1.",
}

# ---------- Filter status penyelesaian ----------
OPSI_STATUS_SEMUA = "Semua status"
OPSI_STATUS_BELUM = "Belum selesai"
OPSI_STATUS_SELESAI = "Selesai"
OPSI_STATUS = [OPSI_STATUS_SEMUA, OPSI_STATUS_BELUM, OPSI_STATUS_SELESAI]
KONDISI_STATUS = {
    OPSI_STATUS_BELUM: "COALESCE(status_penyelesaian, 0) <> 1",
    OPSI_STATUS_SELESAI: "COALESCE(status_penyelesaian, 0) = 1",
}

# ---------- CSV monitoring (update dari DB) ----------
MONITORING_DIR = BASE_DIR / "monitoring_ntb"  # ubah kalau foldernya di tempat lain
TEKS_SUDAH = "Sudah ditindaklanjut"
TEKS_BELUM = "Belum ditindaklanjut"
KOLOM_TINDAK = "Tindak lanjut"
KOLOM_KET = "Keterangan"
KOLOM_KRITERIA = "Kriteria"
KOLOM_OUTPUT_MONEV = [
    "assignment_id",
    "level_2_code",
    "kategori",
    "nama_usaha",
    KOLOM_KRITERIA,
    KOLOM_TINDAK,
    KOLOM_KET,
]
ENCODING_KANDIDAT = ("utf-8-sig", "cp1252", "latin-1")  # latin-1 selalu berhasil

# ---------- Cross tabel ----------
TGL_MULAI = date(2026, 10, 6)
TGL_AKHIR = date(2026, 10, 14)
DAFTAR_TGL = [
    TGL_MULAI + timedelta(days=i) for i in range((TGL_AKHIR - TGL_MULAI).days + 1)
]
OPSI_HITUNG_SELESAI = "Yang diselesaikan (status = selesai)"
OPSI_HITUNG_UPDATE = "Semua yang diupdate (selesai maupun hanya keterangan)"

# ---------- Tampilan tabel data rinci ----------
# Kolom gabungan: nama -> [(kolom asli, label, jenis)], jenis: str / kbli / int / real
KOLOM_GABUNG = {
    "detail_usaha": [
        ("kategori", "kategori", "str"),
        ("kode_kbli", "kbli", "kbli"),
        ("reklasifikasi_skala_usaha", "skala", "str"),
        ("metrik_tenaga_kerja", "tenaga_kerja", "int"),
    ],
    "pengeluaran": [
        ("metrik_upah_gaji", "Upah/gaji", "int"),
        ("metrik_biaya_produksi", "Produksi", "int"),
        ("metrik_biaya_pembelian", "Pembelian", "int"),
        ("metrik_biaya_operasional", "Operasional", "int"),
        ("metrik_biaya_non_operasional", "Non operasional", "int"),
        ("metrik_total_pengeluaran", "Total", "int"),
    ],
    "pendapatan": [
        ("metrik_pendapatan_barang_jasa", "Pendapatan barang/jasa", "int"),
        ("metrik_pendapatan_lainnya", "Pendapatan lainnya", "int"),
        ("metrik_total_aset", "Total aset", "int"),
    ],
    "perhitungan": [
        ("metrik_output", "Output", "int"),
        ("metrik_nilai_tambah", "Nilai tambah", "int"),
        (
            "rasio_biaya_pembelian_barang_terhadap_omzet",
            "Rasio pembelian/omzet",
            "real",
        ),
        ("rasio_ntb", "Rasio NTB", "real"),
        ("produktivitas", "Produktivitas", "real"),
    ],
    "flag": [
        ("flag_rasio_1_output_aset", "flag_1_output_aset", "int"),
        ("flag_rasio_2_upah_ntb", "flag_2_upah_ntb", "int"),
        ("flag_rasio_3_ntb_output", "flag_3_ntb_output", "int"),
        ("flag_kategori_P_dan_U", "flag_P_U", "int"),
    ],
}

# Hanya kolom yang benar-benar ditampilkan yang diambil dari database
KOLOM_TAMPIL = [
    "link",
    "assignment_id",
    "level_2_code",
    "nama_usaha",
    *[c for items in KOLOM_GABUNG.values() for c, _, _ in items],
    "status_penyelesaian",
    "keterangan",
    "update_at",
]

HEADER = [
    "assignment_id",
    "wilayah",
    "nama_usaha",
    *KOLOM_GABUNG,
    "selesai",
    "keterangan",
]
LEBAR = [1.6, 1.0, 2, 2, 2.4, 2.2, 2.6, 2, 0.9, 2.6]

REKAP_SQL = """
    SELECT {grp},
           SUM(COALESCE(status_penyelesaian, 0) <> 1) AS belum,
           SUM(COALESCE(status_penyelesaian, 0) = 1)  AS selesai,
           COUNT(*)                                   AS total
    FROM {tabel} {where}
    GROUP BY {grp}
    ORDER BY {grp}
"""


# ============================================================
# Database
# ============================================================
def connect():
    return closing(sqlite3.connect(DB_PATH, timeout=30))


def query(sql, params=()):
    with connect() as con:
        return pd.read_sql_query(sql, con, params=params)


@st.cache_resource
def siapkan_db():
    """Sekali per sesi server: index + kolom update_at (jika belum ada)."""
    with connect() as con:
        for kol in KOLOM_FILTER:
            con.execute(f"CREATE INDEX IF NOT EXISTS idx_{kol} ON {TABEL}({kol})")
        kolom_ada = [r[1] for r in con.execute(f"PRAGMA table_info({TABEL})")]
        if "update_at" not in kolom_ada:
            con.execute(f"ALTER TABLE {TABEL} ADD COLUMN update_at TEXT")
        con.commit()


siapkan_db()


def bangun_where(
    filter_kolom=None, cek=OPSI_SEMUA_DATA, cari="", status=OPSI_STATUS_SEMUA
):
    """
    filter_kolom: dict {kolom: nilai}. Kolom dengan nilai None diabaikan.
    cari: teks pencarian pada nama_usaha atau assignment_id (LIKE, tidak peka huruf).
    """
    syarat, params = [], []
    for kolom, nilai in dict(filter_kolom or {}).items():
        assert kolom in KOLOM_FILTER, f"Kolom tidak diizinkan: {kolom}"
        if nilai is not None:
            syarat.append(f"{kolom} = ?")
            params.append(nilai)
    if cek in KONDISI_CEK:
        syarat.append(KONDISI_CEK[cek])
    if status in KONDISI_STATUS:
        syarat.append(KONDISI_STATUS[status])

    cari = (cari or "").strip()
    if cari:
        # Escape karakter khusus LIKE agar % dan _ dicari apa adanya
        pola = (
            "%"
            + cari.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")
            + "%"
        )
        syarat.append(
            "(nama_usaha LIKE ? ESCAPE '\\' OR assignment_id LIKE ? ESCAPE '\\')"
        )
        params += [pola, pola]

    where = f"WHERE {' AND '.join(syarat)}" if syarat else ""
    return where, params


def _tuple(d):
    """dict -> tuple terurut (hashable, aman untuk st.cache_data)."""
    return tuple(sorted(d.items()))


@st.cache_data(show_spinner=False)
def get_pilihan(kolom):
    assert kolom in KOLOM_FILTER
    d = query(
        f"SELECT DISTINCT {kolom} AS v FROM {TABEL} WHERE {kolom} IS NOT NULL ORDER BY 1"
    )
    return d["v"].tolist()


@st.cache_data(show_spinner=False)
def get_rekap(grp, filter_kolom=(), cek=OPSI_SEMUA_DATA, status=OPSI_STATUS_SEMUA):
    assert grp in KOLOM_FILTER
    where, params = bangun_where(dict(filter_kolom), cek, "", status)
    return query(REKAP_SQL.format(grp=grp, tabel=TABEL, where=where), params)


@st.cache_data(show_spinner=False)
def get_ringkas(filter_kolom, cek=OPSI_SEMUA_DATA, cari="", status=OPSI_STATUS_SEMUA):
    where, params = bangun_where(dict(filter_kolom), cek, cari, status)
    r = query(
        f"SELECT COUNT(*) AS total, "
        f"COALESCE(SUM(COALESCE(status_penyelesaian, 0) = 1), 0) AS selesai "
        f"FROM {TABEL} {where}",
        params,
    ).iloc[0]
    return int(r["total"]), int(r["selesai"])


@st.cache_data(show_spinner=False)
def get_halaman(filter_kolom, cek, cari, status, ukuran, hal):
    """Hanya satu halaman (LIMIT/OFFSET). rowid ikut diambil sebagai ID unik baris."""
    where, params = bangun_where(dict(filter_kolom), cek, cari, status)
    d = query(
        f"SELECT rowid AS rid, {', '.join(KOLOM_TAMPIL)} FROM {TABEL} {where} "
        f"ORDER BY rowid LIMIT ? OFFSET ?",
        [*params, ukuran, (hal - 1) * ukuran],
    )
    d["status_penyelesaian"] = (
        d["status_penyelesaian"].fillna(0).astype(int).astype(bool)
    )
    d["keterangan"] = d["keterangan"].fillna("")
    return d.to_dict("records")


@st.cache_data(show_spinner=False)
def get_crosstab(filter_kolom=(), cek=OPSI_SEMUA_DATA, hitung=OPSI_HITUNG_SELESAI):
    where, params_where = bangun_where(dict(filter_kolom), cek)

    selesai = "COALESCE(status_penyelesaian, 0) = 1"
    syarat_tgl = f"{selesai} AND " if hitung == OPSI_HITUNG_SELESAI else ""

    kolom_tgl, params_tgl = [], []
    for t in DAFTAR_TGL:
        kolom_tgl.append(
            f"SUM(CASE WHEN {syarat_tgl}date(update_at) = ? THEN 1 ELSE 0 END) "
            f'AS "{t:%d-%m}"'
        )
        params_tgl.append(t.isoformat())

    sql = f"""
        SELECT level_2_code AS wilayah,
               kategori,
               COUNT(*) AS jumlah_data,
               {', '.join(kolom_tgl)},
               SUM(CASE WHEN {selesai} THEN 1 ELSE 0 END) AS total_selesai,
               SUM(CASE WHEN {selesai} THEN 0 ELSE 1 END) AS total_belum
        FROM {TABEL} {where}
        GROUP BY level_2_code, kategori
        ORDER BY level_2_code, kategori
    """
    # Urutan placeholder di SQL: kolom tanggal dulu, baru WHERE
    return query(sql, [*params_tgl, *params_where])


def get_csv(filter_kolom, cek, cari, status):
    """Tidak di-cache; hanya dipanggil saat tombol 'Siapkan CSV' ditekan."""
    where, params = bangun_where(dict(filter_kolom), cek, cari, status)
    d = query(
        f"SELECT * FROM {TABEL} {where} ORDER BY assignment_id, nama_usaha, rowid",
        params,
    )
    d = d.drop(columns=["sumber_file"], errors="ignore")
    return d.to_csv(index=False).encode("utf-8-sig")


def simpan(rid, key_status, key_ket):
    """Dipanggil otomatis (on_change). UPDATE tepat satu baris lewat rowid."""
    status = int(bool(st.session_state[key_status]))
    ket = st.session_state[key_ket] or ""
    with connect() as con:
        # +7 jam = WIB (SQLite menyimpan waktu UTC secara default)
        con.execute(
            f"UPDATE {TABEL} SET status_penyelesaian = ?, keterangan = ?, "
            f"update_at = datetime('now', '+7 hours') WHERE rowid = ?",
            (status, ket, rid),
        )
        con.commit()
    # Bersihkan hanya cache yang terpengaruh (get_pilihan tidak berubah)
    get_halaman.clear()
    get_ringkas.clear()
    get_rekap.clear()
    get_crosstab.clear()
    st.toast("Tersimpan", icon="✅")


# ============================================================
# CSV monitoring (update dari DB)
# ============================================================
def _baca_teks(path):
    """Baca file dengan encoding pertama yang berhasil. Return (teks, encoding)."""
    mentah = path.read_bytes()
    for enc in ENCODING_KANDIDAT:
        try:
            return mentah.decode(enc), enc
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Encoding {path.name} tidak dikenali")


def _norm(s):
    """Normalisasi kunci pencocokan: tanpa spasi tepi, tidak peka huruf besar/kecil."""
    return s.fillna("").astype(str).str.strip().str.casefold()


def get_excel_monitoring(wilayah):
    """
    Baca monitoring_ntb/{wilayah}.csv, isi 'Tindak lanjut' dan 'Keterangan' dari DB
    (dicocokkan lewat assignment_id + nama_usaha), ambil kolom KOLOM_OUTPUT_MONEV,
    lalu tulis ke Excel. Return (bytes_xlsx, total_baris, baris_cocok),
    atau None bila file CSV tidak ada.
    """
    path = MONITORING_DIR / f"{wilayah}.csv"
    if not path.exists():
        return None

    teks, _ = _baca_teks(path)
    header = teks.splitlines()[0]
    sep = max([",", ";", "\t"], key=header.count)  # deteksi pemisah CSV sumber

    d = pd.read_csv(io.StringIO(teks), dtype=str, sep=sep, keep_default_na=False)
    d.columns = d.columns.str.strip()

    for wajib in ("assignment_id", "nama_usaha"):
        if wajib not in d.columns:
            raise ValueError(
                f"Kolom '{wajib}' tidak ada di {path.name}. Kolom: {list(d.columns)}"
            )

    # Data DB untuk wilayah ini saja
    db = query(
        f"SELECT assignment_id, nama_usaha, level_2_code, kategori, "
        f"status_penyelesaian, keterangan FROM {TABEL} WHERE level_2_code = ?",
        (wilayah,),
    )
    db["_a"] = _norm(db["assignment_id"])
    db["_n"] = _norm(db["nama_usaha"])
    db["_status"] = db["status_penyelesaian"].fillna(0).astype(int)
    db["_ket"] = db["keterangan"].fillna("").astype(str).str.strip()

    # Kunci kembar di DB: status tertinggi, keterangan tidak kosong digabung
    agg = db.groupby(["_a", "_n"], as_index=False).agg(
        _status=("_status", "max"),
        _ket=("_ket", lambda s: " | ".join(dict.fromkeys(x for x in s if x))),
        _wil=("level_2_code", "first"),
        _kat=("kategori", "first"),
    )

    d["_a"] = _norm(d["assignment_id"])
    d["_n"] = _norm(d["nama_usaha"])
    m = d.merge(agg, on=["_a", "_n"], how="left")

    ada = m["_status"].notna()
    for kol in (KOLOM_TINDAK, KOLOM_KET):
        if kol not in m.columns:
            m[kol] = ""
    m.loc[ada, KOLOM_TINDAK] = m.loc[ada, "_status"].map(
        lambda s: TEKS_SUDAH if s == 1 else TEKS_BELUM
    )
    m.loc[ada, KOLOM_KET] = m.loc[ada, "_ket"]

    # Kolom yang tidak ada di CSV: wilayah/kategori diambil dari DB, sisanya kosong
    for kol, cadangan in (("level_2_code", "_wil"), ("kategori", "_kat")):
        if kol not in m.columns:
            m[kol] = m[cadangan].fillna("")
    if KOLOM_KRITERIA not in m.columns:
        m[KOLOM_KRITERIA] = ""

    hasil = m[KOLOM_OUTPUT_MONEV].reset_index(drop=True)

    # ---- Tulis ke Excel ----
    buf = io.BytesIO()
    with pd.ExcelWriter(
        buf, engine="xlsxwriter", engine_kwargs={"options": {"strings_to_urls": False}}
    ) as writer:
        hasil.to_excel(writer, sheet_name=str(wilayah), index=False)
        ws, wb = writer.sheets[str(wilayah)], writer.book
        fmt_header = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
        fmt_teks = wb.add_format({"num_format": "@"})  # kode tetap teks
        fmt_wrap = wb.add_format({"text_wrap": True, "valign": "top"})
        teks_kolom = {"assignment_id", "level_2_code"}

        for j, col in enumerate(hasil.columns):
            ws.write(0, j, col, fmt_header)
            lebar = max(
                len(str(col)),
                hasil[col].head(1000).astype(str).str.len().max() if len(hasil) else 0,
            )
            fmt = (
                fmt_teks
                if col in teks_kolom
                else (fmt_wrap if col == KOLOM_KET else None)
            )
            ws.set_column(j, j, min(lebar + 2, 50), fmt)
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, max(len(hasil), 1), len(hasil.columns) - 1)

    return buf.getvalue(), len(m), int(ada.sum())


def unduh_monev(utama, filter_kolom):
    wilayah = dict(filter_kolom).get("level_2_code")

    if st.button(
        "Siapkan Excel monitoring (update dari DB)",
        key=f"monev_btn_{utama}",
        disabled=wilayah is None,
        help=f"Pilih kode wilayah dulu. Sumber data: {MONITORING_DIR.name}/<kode>.csv",
    ):
        try:
            hasil = get_excel_monitoring(wilayah)
        except Exception as e:
            st.error(f"Gagal memproses CSV: {e}")
            return

        if hasil is None:
            st.error(f"File tidak ditemukan: {MONITORING_DIR / f'{wilayah}.csv'}")
            return

        data, total, cocok = hasil
        if cocok < total:
            st.warning(
                f"{total - cocok:,} dari {total:,} baris tidak ditemukan di DB "
                "(assignment_id + nama_usaha tidak cocok), kolom Tindak lanjut "
                "dan Keterangan dibiarkan seperti CSV asli."
            )
        st.download_button(
            f"Download {wilayah}.xlsx ({total:,} baris)",
            data,
            file_name=f"{wilayah}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"monev_dl_{utama}",
        )


# ============================================================
# Format & tampilan
# ============================================================
def _angka(v, desimal):
    """Format Indonesia: titik ribuan, koma desimal."""
    s = f"{float(v):,.{desimal}f}"
    return s.replace(",", "#").replace(".", ",").replace("#", ".")


FORMAT = {
    "str": lambda v: escape(str(v)),
    "kbli": lambda v: escape(str(v).strip().removesuffix(".0").zfill(5)),
    "int": lambda v: _angka(v, 0),
    "real": lambda v: _angka(v, 4),
}


def fmt(v, jenis):
    if v is None or pd.isna(v):
        return "-"
    try:
        return FORMAT[jenis](v)
    except (TypeError, ValueError):
        return escape(str(v))


def sel(isi):
    return f'<div style="font-size:13px;line-height:1.5">{isi}</div>'


def html_gabung(row, items):
    return "<br>".join(
        f"{escape(label)}: {fmt(row[col], jenis)}" for col, label, jenis in items
    )


def teks_atau_strip(v):
    return "-" if pd.isna(v) or v == "" else escape(str(v))


def tampil_baris(rows, kunci):
    head = st.columns(LEBAR)
    for c, nama in zip(head, HEADER):
        c.markdown(f"**{nama}**")
    st.divider()

    for r in rows:
        rid = int(r["rid"])  # unik per baris, dipakai untuk key & UPDATE
        aid = str(r["assignment_id"])
        ks, kk = f"status_{kunci}_{rid}", f"ket_{kunci}_{rid}"
        cols = st.columns(LEBAR, vertical_alignment="top")

        # assignment_id sekaligus link ke FASIH
        if pd.notna(r["link"]):
            href = escape(str(r["link"]), quote=True)
            cols[0].html(sel(f'<a href="{href}" target="_blank">{escape(aid)}</a>'))
        else:
            cols[0].html(sel(escape(aid)))

        cols[1].html(sel(teks_atau_strip(r["level_2_code"])))
        cols[2].html(sel(teks_atau_strip(r["nama_usaha"])))

        # kolom gabungan (detail, pengeluaran, pendapatan, perhitungan, flag)
        for col, items in zip(cols[3:8], KOLOM_GABUNG.values()):
            col.html(sel(html_gabung(r, items)))

        cols[8].checkbox(
            "selesai",
            value=r["status_penyelesaian"],
            key=ks,
            label_visibility="collapsed",
            on_change=simpan,
            args=(rid, ks, kk),
        )
        cols[9].text_area(
            "keterangan",
            value=r["keterangan"],
            key=kk,
            height=68,
            label_visibility="collapsed",
            on_change=simpan,
            args=(rid, ks, kk),
        )
        upd = r["update_at"]
        cols[9].caption(
            f"Update: {upd} WIB" if pd.notna(upd) and upd else "Belum pernah diupdate"
        )
        st.divider()


def tampil_raw(kunci, filter_kolom, cek, cari, status, judul, nama_file):
    total, selesai = get_ringkas(filter_kolom, cek, cari, status)

    st.subheader(judul)
    c1, c2, c3 = st.columns(3)
    c1.metric("Jumlah usaha", f"{total:,}")
    c2.metric("Selesai", f"{selesai:,}")
    c3.metric("Belum selesai", f"{total - selesai:,}")

    if total == 0:
        st.info("Tidak ada data yang cocok dengan filter / pencarian.")
        return

    # Pagination: widget per baris itu berat, jadi halaman dibatasi
    p1, p2 = st.columns([1, 3])
    ukuran = p1.selectbox(
        "Baris per halaman", [10, 25, 50, 100], index=1, key=f"ukuran_{kunci}"
    )
    total_hal = max(1, -(-total // ukuran))
    # Key memuat semua filter agar nomor halaman reset saat berubah
    sidik = "_".join(f"{v}" for _, v in filter_kolom)
    hal = p2.number_input(
        f"Halaman (dari {total_hal})",
        min_value=1,
        max_value=total_hal,
        value=1,
        step=1,
        key=(
            f"hal_{kunci}_{sidik}_{OPSI_CEK.index(cek)}_"
            f"{OPSI_STATUS.index(status)}_{cari}_{ukuran}"
        ),
    )

    awal = (hal - 1) * ukuran + 1
    st.caption(
        f"Menampilkan baris {min(awal, total):,} - {min(hal * ukuran, total):,} dari {total:,}. "
        "Centang status atau isi keterangan (klik di luar kotak / Ctrl+Enter) "
        "untuk menyimpan otomatis."
    )
    tampil_baris(get_halaman(filter_kolom, cek, cari, status, ukuran, hal), kunci)

    # CSV hanya dibuat saat diminta, bukan di setiap rerun
    if st.button("Siapkan CSV", key=f"csv_{kunci}"):
        st.download_button(
            "Download CSV",
            get_csv(filter_kolom, cek, cari, status),
            file_name=nama_file.replace("/", "_"),
            mime="text/csv",
            key=f"dl_{kunci}",
        )


# ============================================================
# Halaman
# ============================================================
def halaman_data(utama, kedua, judul_halaman):
    """
    utama : kolom filter utama (menentukan pengelompokan rekap)
    kedua : kolom filter silang (opsional, default Semua)
    """
    st.title(judul_halaman)

    col1, col2, col3 = st.columns(3)
    pilih_utama = col1.selectbox(
        f"Filter {LABEL[utama]}", [SEMUA] + get_pilihan(utama), key=f"f_{utama}_utama"
    )
    pilih_kedua = col2.selectbox(
        f"Filter {LABEL[kedua]}",
        [SEMUA] + get_pilihan(kedua),
        key=f"f_{kedua}_silang_{utama}",
    )
    cek = col3.selectbox("Filter pengecekan", OPSI_CEK, key=f"cek_{utama}")

    cs1, cs2 = st.columns([3, 1])
    cari = cs1.text_input(
        "Cari nama usaha / assignment_id",
        placeholder="Ketik lalu tekan Enter...",
        key=f"cari_{utama}",
    ).strip()
    status = cs2.selectbox("Filter status", OPSI_STATUS, key=f"status_{utama}")

    # Hanya filter yang dipilih (bukan Semua) yang ikut ke query
    aktif = {}
    if pilih_utama != SEMUA:
        aktif[utama] = pilih_utama
    if pilih_kedua != SEMUA:
        aktif[kedua] = pilih_kedua
    filter_kolom = _tuple(aktif)

    # Rekap hanya jika filter utama Semua DAN tidak sedang mencari
    if pilih_utama == SEMUA and not cari:
        st.subheader(f"Rekap penyelesaian per {LABEL[utama]}")
        agg = get_rekap(utama, filter_kolom, cek, status)

        info = []
        if cek in DESKRIPSI_CEK:
            info.append(DESKRIPSI_CEK[cek])
        if pilih_kedua != SEMUA:
            info.append(f"Terfilter {LABEL[kedua]} = {pilih_kedua}.")
        if status != OPSI_STATUS_SEMUA:
            info.append(f"Status: {status}.")
        info.append(f"{len(agg):,} {LABEL[utama]}. Belum = status 0 atau kosong.")
        info.append(
            f"Pilih {LABEL[utama]} atau isi pencarian untuk melihat data rinci."
        )
        st.caption(" ".join(info))

        m1, m2, m3 = st.columns(3)
        m1.metric("Total", f"{int(agg['total'].sum()):,}")
        m2.metric("Belum", f"{int(agg['belum'].sum()):,}")
        m3.metric("Selesai", f"{int(agg['selesai'].sum()):,}")
        st.dataframe(agg, use_container_width=True, hide_index=True)
    else:
        judul = (
            f"Data rinci {pilih_utama}" if pilih_utama != SEMUA else "Hasil pencarian"
        )
        if pilih_kedua != SEMUA:
            judul += f" | {LABEL[kedua]}: {pilih_kedua}"
        if cari:
            judul += f' | cari: "{cari}"'
        if status != OPSI_STATUS_SEMUA:
            judul += f" | {status}"
        if cek in KONDISI_CEK:
            judul += f" ({cek})"

        nama_file = f"{utama}_{pilih_utama}" if pilih_utama != SEMUA else "pencarian"
        if pilih_kedua != SEMUA:
            nama_file += f"_{kedua}_{pilih_kedua}"
        if status != OPSI_STATUS_SEMUA:
            nama_file += f"_{status.replace(' ', '_')}"
        if cari:
            nama_file += f"_{cari}"
        tampil_raw(utama, filter_kolom, cek, cari, status, judul, nama_file + ".csv")

    # Tombol CSV monitoring: aktif kalau kode wilayah sudah dipilih
    unduh_monev(utama, filter_kolom)


def halaman_wilayah():
    halaman_data("level_2_code", "kategori", "Data By Wilayah")


def halaman_kategori():
    halaman_data("kategori", "level_2_code", "Data By Kategori")


def halaman_crosstab():
    st.title("Cross Tabel Wilayah x Kategori")

    c1, c2, c3 = st.columns(3)
    pilih_wil = c1.selectbox(
        "Filter kode wilayah", [SEMUA] + get_pilihan("level_2_code"), key="wil_crosstab"
    )
    pilih_kat = c2.selectbox(
        "Filter kategori", [SEMUA] + get_pilihan("kategori"), key="kat_crosstab"
    )
    cek = c3.selectbox("Filter pengecekan", OPSI_CEK, key="cek_crosstab")

    hitung = st.selectbox(
        "Yang dihitung per tanggal",
        [OPSI_HITUNG_SELESAI, OPSI_HITUNG_UPDATE],
        key="hitung_crosstab",
    )

    aktif = {}
    if pilih_wil != SEMUA:
        aktif["level_2_code"] = pilih_wil
    if pilih_kat != SEMUA:
        aktif["kategori"] = pilih_kat

    df = get_crosstab(_tuple(aktif), cek, hitung)
    if df.empty:
        st.info("Tidak ada data yang cocok dengan filter.")
        return

    # Baris total di paling bawah
    kolom_angka = [c for c in df.columns if c not in ("wilayah", "kategori")]
    total = df[kolom_angka].sum()
    df_total = pd.DataFrame([{"wilayah": "TOTAL", "kategori": "", **total.to_dict()}])
    tampil = pd.concat([df, df_total], ignore_index=True)

    m1, m2, m3 = st.columns(3)
    m1.metric("Jumlah data", f"{int(total['jumlah_data']):,}")
    m2.metric("Total selesai", f"{int(total['total_selesai']):,}")
    m3.metric("Total belum", f"{int(total['total_belum']):,}")

    st.caption(
        f"Kolom tanggal ({TGL_MULAI:%d-%m} s/d {TGL_AKHIR:%d-%m}-{TGL_AKHIR.year}) "
        "dihitung dari tanggal update_at. Total selesai/belum mencakup semua "
        "tanggal, termasuk baris yang update_at-nya kosong."
    )
    st.dataframe(tampil, use_container_width=True, hide_index=True)

    nama_file = "crosstab"
    if pilih_wil != SEMUA:
        nama_file += f"_{pilih_wil}"
    if pilih_kat != SEMUA:
        nama_file += f"_{pilih_kat}"
    st.download_button(
        "Download CSV",
        tampil.to_csv(index=False).encode("utf-8-sig"),
        file_name=nama_file.replace("/", "_") + ".csv",
        mime="text/csv",
        key="dl_crosstab",
    )


pg = st.navigation(
    [
        st.Page(
            halaman_wilayah, title="By Wilayah", icon=":material/map:", default=True
        ),
        st.Page(halaman_kategori, title="By Kategori", icon=":material/category:"),
        st.Page(halaman_crosstab, title="Cross Tabel", icon=":material/table_chart:"),
    ]
)
pg.run()
