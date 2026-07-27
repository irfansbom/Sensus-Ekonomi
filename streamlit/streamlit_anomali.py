import streamlit as st
import sqlite3
import pandas as pd
from pathlib import Path
import io
import re

# Layout wide agar memanfaatkan lebar layar laptop standar
st.set_page_config(page_title="Rekap Anomali Pertanggal", layout="wide")
st.title("Rekap Anomali Pertanggal")
st.info(
    "📊 Tabel ini diambil dari dashboard dan **diperbarui secara berkala**. "
    "Tujuannya untuk memantau progress anomali pertanggal, "
    "sehingga memudahkan monitoring dan penyelesaian anomali yang realtime."
)

# Path absolut berdasarkan lokasi file script ini
BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR.parent / "SQLLITE" / "rekap_anomali_pendataan.db"


@st.cache_data(ttl=60)  # cache refresh tiap 60 detik
def load_data(db_path):
    conn = sqlite3.connect(db_path)
    df = pd.read_sql_query("SELECT * FROM rekap_anomali_pendataan", conn)
    conn.close()
    return df


df = load_data(str(DB_PATH))

# ==== FILTER & SEARCH DI SIDEBAR ====
st.sidebar.header("Filter & Pencarian")

# Filter dropdown berdasarkan kd_kab
kab_list = sorted(df["kode_kabupaten"].dropna().unique().tolist())
selected_kab = st.sidebar.multiselect("Filter Kode Kabupaten (kd_kab)", kab_list)

type_list = sorted(df["source_type"].dropna().unique().tolist())
selected_type = st.sidebar.multiselect("Filter Tipe Anomali", type_list)

anomali_no = sorted(df["anomali_no"].dropna().unique().tolist())
selected_anomali_no = st.sidebar.multiselect(
    "Filter Nomor Anomali", anomali_no
)

# ==== TERAPKAN FILTER ====
filtered_df = df.copy()

if selected_kab:
    filtered_df = filtered_df[filtered_df["kode_kabupaten"].isin(selected_kab)]
if selected_type:
    filtered_df = filtered_df[filtered_df["source_type"].isin(selected_type)]
if selected_anomali_no:
    filtered_df = filtered_df[filtered_df["anomali_no"].isin(selected_anomali_no)]


filtered_df = filtered_df.sort_values(by="level_6_code", ascending=True).reset_index(
    drop=True
)

# ==== DEFINISIKAN SEMUA KOLOM ====
all_columns = df.columns.tolist()

# ==== DETEKSI KOLOM YANG NAMANYA TANGGAL (format dd-mm-yyyy) ====
date_col_pattern = re.compile(r"^\d{2}-\d{2}-\d{4}$")

date_columns = [c for c in all_columns if date_col_pattern.match(c)]
date_columns_sorted = sorted(
    date_columns, key=lambda x: pd.to_datetime(x, format="%d-%m-%Y")
)

st.sidebar.subheader("Pilih Kolom Tanggal")

# Default: tampilkan cuma tanggal PALING BARU
default_date_cols = date_columns_sorted[-1:] if date_columns_sorted else []
# ==== FILTER ISI KOLOM TANGGAL (mirip filter Excel per kolom) ====
st.sidebar.subheader("Filter Isian Kolom Tanggal")

if date_columns_sorted:
    filter_date_col = st.sidebar.selectbox(
        "Pilih kolom tanggal untuk difilter isinya",
        options=["(Tidak difilter)"] + date_columns_sorted,
    )
else:
    filter_date_col = "(Tidak difilter)"

if filter_date_col != "(Tidak difilter)":
    # Ambil nilai unik di kolom tsb, tangani NaN/kosong terpisah
    raw_values = filtered_df[filter_date_col]
    unique_values = raw_values.dropna().unique().tolist()
    unique_values = [v for v in unique_values if str(v).strip() != ""]
    unique_values = sorted(unique_values, key=lambda x: str(x))

    has_blank = (
        raw_values.isna().any() or (raw_values.astype(str).str.strip() == "").any()
    )

    options = unique_values.copy()
    if has_blank:
        options = ["(Kosong/Blank)"] + options

    selected_values = st.sidebar.multiselect(
        f"Tampilkan baris dengan isian '{filter_date_col}'",
        options=options,
        default=[],
    )

    if selected_values:
        mask = pd.Series(False, index=filtered_df.index)

        if "(Kosong/Blank)" in selected_values:
            mask |= filtered_df[filter_date_col].isna() | (
                filtered_df[filter_date_col].astype(str).str.strip() == ""
            )

        other_values = [v for v in selected_values if v != "(Kosong/Blank)"]
        if other_values:
            mask |= filtered_df[filter_date_col].isin(other_values)

        filtered_df = filtered_df[mask]

display_df = filtered_df

st.write(f"Menampilkan **{len(filtered_df)}** dari **{len(df)}** total baris")
st.dataframe(
    display_df,
    use_container_width=True,
    height=650,
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
