import sqlite3
from contextlib import closing
from html import escape
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Monitoring Usaha", layout="wide")

# ============================================================
# Konfigurasi
# ============================================================
DB_PATH = (
    Path(__file__).resolve().parent
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

# ---------- Filter pengecekan ----------
AMBANG_OUTPUT = 1_000_000_000  # output > 1 miliar
AMBANG_NTB = -5_000_000  # nilai tambah < -5 juta

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
    "nama_usaha",
    *[c for items in KOLOM_GABUNG.values() for c, _, _ in items],
    "status_penyelesaian",
    "keterangan",
]

HEADER = ["assignment_id", "nama_usaha", *KOLOM_GABUNG, "selesai", "keterangan"]
LEBAR = [1.6, 2, 2, 2.4, 2.2, 2.6, 2, 0.9, 2.6]

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
    """Sekali per sesi server: index agar filter & UPDATE per assignment_id cepat."""
    with connect() as con:
        for kol in (*KOLOM_FILTER, "assignment_id"):
            con.execute(f"CREATE INDEX IF NOT EXISTS idx_{kol} ON {TABEL}({kol})")
        con.commit()


siapkan_db()


def bangun_where(kolom, nilai=None, cek=OPSI_SEMUA_DATA):
    assert kolom in KOLOM_FILTER, f"Kolom tidak diizinkan: {kolom}"
    syarat, params = [], []
    if nilai is not None:
        syarat.append(f"{kolom} = ?")
        params.append(nilai)
    if cek in KONDISI_CEK:
        syarat.append(KONDISI_CEK[cek])
    where = f"WHERE {' AND '.join(syarat)}" if syarat else ""
    return where, params


@st.cache_data(show_spinner=False)
def get_pilihan(kolom):
    assert kolom in KOLOM_FILTER
    d = query(
        f"SELECT DISTINCT {kolom} AS v FROM {TABEL} WHERE {kolom} IS NOT NULL ORDER BY 1"
    )
    return d["v"].tolist()


@st.cache_data(show_spinner=False)
def get_rekap(kolom, cek=OPSI_SEMUA_DATA):
    where, _ = bangun_where(kolom, None, cek)
    return query(REKAP_SQL.format(grp=kolom, tabel=TABEL, where=where))


@st.cache_data(show_spinner=False)
def get_ringkas(kolom, nilai, cek=OPSI_SEMUA_DATA):
    where, params = bangun_where(kolom, nilai, cek)
    r = query(
        f"SELECT COUNT(*) AS total, "
        f"COALESCE(SUM(COALESCE(status_penyelesaian, 0) = 1), 0) AS selesai "
        f"FROM {TABEL} {where}",
        params,
    ).iloc[0]
    return int(r["total"]), int(r["selesai"])


@st.cache_data(show_spinner=False)
def get_halaman(kolom, nilai, cek, ukuran, hal):
    """Hanya satu halaman (LIMIT/OFFSET), hanya kolom yang ditampilkan."""
    where, params = bangun_where(kolom, nilai, cek)
    d = query(
        f"SELECT {', '.join(KOLOM_TAMPIL)} FROM {TABEL} {where} "
        f"ORDER BY rowid LIMIT ? OFFSET ?",
        [*params, ukuran, (hal - 1) * ukuran],
    )
    d["status_penyelesaian"] = (
        d["status_penyelesaian"].fillna(0).astype(int).astype(bool)
    )
    d["keterangan"] = d["keterangan"].fillna("")
    return d.to_dict("records")


def get_csv(kolom, nilai, cek):
    """Tidak di-cache; hanya dipanggil saat tombol 'Siapkan CSV' ditekan."""
    where, params = bangun_where(kolom, nilai, cek)
    d = query(f"SELECT * FROM {TABEL} {where}", params)
    d = d.drop(columns=["sumber_file"], errors="ignore")
    return d.to_csv(index=False).encode("utf-8-sig")


def simpan(aid, key_status, key_ket):
    """Dipanggil otomatis (on_change) saat checkbox / keterangan berubah."""
    status = int(bool(st.session_state[key_status]))
    ket = st.session_state[key_ket] or ""
    with connect() as con:
        con.execute(
            f"UPDATE {TABEL} SET status_penyelesaian = ?, keterangan = ? "
            f"WHERE assignment_id = ?",
            (status, ket, aid),
        )
        con.commit()
    # Bersihkan hanya cache yang terpengaruh (get_pilihan tidak berubah)
    get_halaman.clear()
    get_ringkas.clear()
    get_rekap.clear()
    st.toast(f"Tersimpan: {aid}", icon="✅")


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


def tampil_baris(rows, kolom):
    head = st.columns(LEBAR)
    for c, nama in zip(head, HEADER):
        c.markdown(f"**{nama}**")
    st.divider()

    for r in rows:
        aid = str(r["assignment_id"])
        ks, kk = f"status_{kolom}_{aid}", f"ket_{kolom}_{aid}"
        cols = st.columns(LEBAR, vertical_alignment="top")

        # assignment_id sekaligus link ke FASIH
        if pd.notna(r["link"]):
            href = escape(str(r["link"]), quote=True)
            cols[0].html(sel(f'<a href="{href}" target="_blank">{escape(aid)}</a>'))
        else:
            cols[0].html(sel(escape(aid)))

        nama = r["nama_usaha"]
        cols[1].html(sel("-" if pd.isna(nama) or nama == "" else escape(str(nama))))

        for col, items in zip(cols[2:7], KOLOM_GABUNG.values()):
            col.html(sel(html_gabung(r, items)))

        cols[7].checkbox(
            "selesai",
            value=r["status_penyelesaian"],
            key=ks,
            label_visibility="collapsed",
            on_change=simpan,
            args=(aid, ks, kk),
        )
        cols[8].text_area(
            "keterangan",
            value=r["keterangan"],
            key=kk,
            height=68,
            label_visibility="collapsed",
            on_change=simpan,
            args=(aid, ks, kk),
        )
        st.divider()


def tampil_raw(kolom, nilai, cek, judul, nama_file):
    total, selesai = get_ringkas(kolom, nilai, cek)

    st.subheader(judul)
    c1, c2, c3 = st.columns(3)
    c1.metric("Jumlah usaha", f"{total:,}")
    c2.metric("Selesai", f"{selesai:,}")
    c3.metric("Belum selesai", f"{total - selesai:,}")

    # Pagination: widget per baris itu berat, jadi halaman dibatasi
    p1, p2 = st.columns([1, 3])
    ukuran = p1.selectbox(
        "Baris per halaman", [10, 25, 50, 100], index=1, key=f"ukuran_{kolom}"
    )
    total_hal = max(1, -(-total // ukuran))
    # Key memuat filter & ukuran agar nomor halaman reset otomatis saat filter berubah
    hal = p2.number_input(
        f"Halaman (dari {total_hal})",
        min_value=1,
        max_value=total_hal,
        value=1,
        step=1,
        key=f"hal_{kolom}_{nilai}_{OPSI_CEK.index(cek)}_{ukuran}",
    )

    awal = (hal - 1) * ukuran + 1
    st.caption(
        f"Menampilkan baris {min(awal, total):,} - {min(hal * ukuran, total):,} dari {total:,}. "
        "Centang status atau isi keterangan (klik di luar kotak / Ctrl+Enter) "
        "untuk menyimpan otomatis."
    )
    tampil_baris(get_halaman(kolom, nilai, cek, ukuran, hal), kolom)

    # CSV hanya dibuat saat diminta, bukan di setiap rerun
    if st.button("Siapkan CSV", key=f"csv_{kolom}"):
        st.download_button(
            "Download CSV",
            get_csv(kolom, nilai, cek),
            file_name=nama_file.replace("/", "_"),
            mime="text/csv",
            key=f"dl_{kolom}",
        )


# ============================================================
# Halaman
# ============================================================
def halaman_data(kolom, judul_halaman, label_filter, pakai_cek):
    st.title(judul_halaman)

    if pakai_cek:
        col1, col2 = st.columns(2)
        terpilih = col1.selectbox(
            f"Filter {label_filter}", [SEMUA] + get_pilihan(kolom), key=f"f_{kolom}"
        )
        cek = col2.selectbox("Filter pengecekan", OPSI_CEK, key=f"cek_{kolom}")
    else:
        terpilih = st.selectbox(
            f"Filter {label_filter}", [SEMUA] + get_pilihan(kolom), key=f"f_{kolom}"
        )
        cek = OPSI_SEMUA_DATA

    if terpilih == SEMUA:
        st.subheader(f"Rekap penyelesaian per {label_filter}")
        agg = get_rekap(kolom, cek)
        if cek in DESKRIPSI_CEK:
            st.caption(f"{DESKRIPSI_CEK[cek]} Belum = status 0 atau kosong.")
        else:
            st.caption(
                f"{len(agg):,} {label_filter}. Belum = status 0 atau kosong. "
                f"Pilih {label_filter} di atas untuk melihat data rinci."
            )
        m1, m2, m3 = st.columns(3)
        m1.metric("Total", f"{int(agg['total'].sum()):,}")
        m2.metric("Belum", f"{int(agg['belum'].sum()):,}")
        m3.metric("Selesai", f"{int(agg['selesai'].sum()):,}")
        st.dataframe(agg, use_container_width=True, hide_index=True)
    else:
        judul = f"Data rinci {terpilih}" + (f" ({cek})" if cek in KONDISI_CEK else "")
        tampil_raw(kolom, terpilih, cek, judul, f"{kolom}_{terpilih}.csv")


def halaman_wilayah():
    halaman_data("level_2_code", "Data By Wilayah", "kode wilayah", pakai_cek=True)


def halaman_kategori():
    halaman_data("kategori", "Data By Kategori", "kategori", pakai_cek=True)


pg = st.navigation(
    [
        st.Page(
            halaman_wilayah, title="By Wilayah", icon=":material/map:", default=True
        ),
        st.Page(halaman_kategori, title="By Kategori", icon=":material/category:"),
    ]
)
pg.run()
