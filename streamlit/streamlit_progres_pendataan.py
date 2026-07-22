import streamlit as st
import sqlite3
import pandas as pd
from pathlib import Path
import io

# Layout wide agar memanfaatkan lebar layar laptop standar
st.set_page_config(page_title="Keberadaan Usaha By Subsls", layout="wide")
st.title("Keberadaan Usaha By Subsls")
st.info(
    "📊 Tabel ini diambil dari dashboard dan **diperbarui secara berkala**. "
    "Tujuannya untuk memantau progres keberadaan usaha per SubSLS, "
    "sehingga memudahkan monitoring dan alokasi petugas (PML/PPL) di lapangan."
)

# Path absolut berdasarkan lokasi file script ini
BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR.parent / "SQLLITE" / "rekap_progress_pendataan.db"


@st.cache_data(ttl=60)  # cache refresh tiap 60 detik
def load_data(db_path):
    conn = sqlite3.connect(db_path)
    df = pd.read_sql_query("SELECT * FROM rekap_progress_pendataan", conn)
    conn.close()
    return df


df = load_data(str(DB_PATH))

# ==== FILTER & SEARCH DI SIDEBAR ====
st.sidebar.header("Filter & Pencarian")

# Filter dropdown berdasarkan kd_kab
kab_list = sorted(df["kd_kab"].dropna().unique().tolist())
selected_kab = st.sidebar.multiselect("Filter Kode Kabupaten (kd_kab)", kab_list)

# Search text untuk beberapa kolom
search_sls = st.sidebar.text_input("Cari Nama SLS")
search_pml = st.sidebar.text_input("Cari PML")
search_ppl = st.sidebar.text_input("Cari PPL")


# ==== PILIH KOLOM YANG DITAMPILKAN ====
# st.sidebar.header("Tampilan Kolom")

all_columns = df.columns.tolist()

# Ganti list ini sesuai kolom yang ingin muncul PERTAMA KALI
# default_columns = [
#     "id_wilayah",
#     "kd_kab",
#     "nama_sls",
#     "PPL",
#     "PML",
#     "jumlah_prelist_usaha",
#     "jumlah_usaha_realisasi",
#     "jumlah_prelist_keluarga",
#     "jumlah_keluarga_realisasi",
#     "last_update",
# ]

# default_columns = [c for c in default_columns if c in all_columns]

# selected_columns = st.sidebar.multiselect(
#     "Pilih kolom yang ditampilkan", options=all_columns, default=default_columns
# )


# ==== TERAPKAN FILTER ====
filtered_df = df.copy()

if selected_kab:
    filtered_df = filtered_df[filtered_df["kd_kab"].isin(selected_kab)]

if search_sls:
    filtered_df = filtered_df[
        filtered_df["nama_sls"].str.contains(search_sls, case=False, na=False)
    ]

if search_pml:
    filtered_df = filtered_df[
        filtered_df["PML"].str.contains(search_pml, case=False, na=False)
    ]

if search_ppl:
    filtered_df = filtered_df[
        filtered_df["PPL"].str.contains(search_ppl, case=False, na=False)
    ]

# ==== TERAPKAN PILIHAN KOLOM (INI YANG KEMARIN KELEWAT) ====
filtered_df = filtered_df.sort_values(by="id_wilayah", ascending=True).reset_index(drop=True)

# ==== TERAPKAN PILIHAN KOLOM ====
# if selected_columns:
#     display_df = filtered_df[selected_columns]
# else:
    # display_df = filtered_df
display_df = filtered_df


# ==== KETERANGAN KOLOM (untuk expander & tooltip) ====
column_descriptions = {
    "jumlah_prelist_usaha": "Prelist UB + Prelist UM + Prelist UMK",
    "jumlah_usaha_realisasi": "Usaha Ditemukan (BKU) + Usaha Baru (BKU)",
    "jumlah_prelist_keluarga": "Jumlah keluarga hasil prelist awal",
    "jumlah_keluarga_realisasi": "Keluarga Ditemukan + Keluarga Baru",
    # tambahkan kolom lain di sini sesuai kebutuhan
}

# ==== TAMPILKAN INFO & TABEL ====
st.write(f"Menampilkan **{len(filtered_df)}** dari **{len(df)}** total baris")


with st.expander("ℹ️ Keterangan Kolom"):
    for col, desc in column_descriptions.items():
        st.markdown(f"- **`{col}`** = {desc}")

# Bangun column_config hanya untuk kolom yang ada keterangannya
column_config = {
    col: st.column_config.Column(help=desc)
    for col, desc in column_descriptions.items()
    if col in display_df.columns
}

st.dataframe(
    display_df,
    use_container_width=True,
    height=600,
    column_config=column_config,
)


# ==== FUNGSI KONVERSI DF KE EXCEL (DI MEMORI) ====
def convert_df_to_excel(dataframe):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="Data")
    return output.getvalue()

# ==== TOMBOL DOWNLOAD CSV & EXCEL ====
col1, col2 = st.columns(2)

with col1:
    csv_data = display_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label="⬇️ Download CSV",
        data=csv_data,
        file_name="keberadaan_usaha_subsls.csv",
        mime="text/csv",
        use_container_width=True,
    )

with col2:
    excel_data = convert_df_to_excel(display_df)
    st.download_button(
        label="⬇️ Download Excel",
        data=excel_data,
        file_name="keberadaan_usaha_subsls.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )