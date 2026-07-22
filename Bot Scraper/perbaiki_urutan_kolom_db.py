"""
Script SEKALI PAKAI untuk merapikan urutan kolom di database SQLite yang
sudah terlanjur dibuat sebelumnya (last_update ketiban di tengah-tengah
kolom tanggal akibat ALTER TABLE ADD COLUMN).

Tidak melakukan scraping apa pun -- cuma baca tabel yang ada, susun ulang
urutan kolomnya, lalu tulis balik ke tabel yang sama.

Setelah dijalankan sekali, script utama (scrap_rekap_anomali.py) yang sudah
dipakai selanjutnya akan otomatis me-replace tabel dengan urutan yang benar
tiap run, jadi script ini tidak perlu dijalankan lagi.
"""

import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd

# ── Sesuaikan 3 baris ini kalau perlu ───────────────────────────────────
DB_PATH = Path("../SQLLITE") / "rekap_anomali_pendataan.db"
TABLE_NAME = "rekap_anomali_pendataan"
FORMAT_TANGGAL = "%d-%m-%Y"

KEY_COLS = [
    "assignment_id",
    "kode_kabupaten",
    "level_6_code",
    "nama_tercantum",
    "source_type",
    "anomali_no",
]
EXTRA_COLS = ["extra_columns"]
LINK_COL = "hyperlink"
LAST_UPDATE_COL = "last_update"


def _is_kolom_tanggal(col: str) -> bool:
    try:
        datetime.strptime(col, FORMAT_TANGGAL)
        return True
    except (ValueError, TypeError):
        return False


def urutkan_kolom(df: pd.DataFrame) -> pd.DataFrame:
    """key -> extra_columns -> hyperlink -> last_update -> tanggal (kronologis)."""
    kolom_tanggal = [c for c in df.columns if _is_kolom_tanggal(c)]
    kolom_tanggal_sorted = sorted(
        kolom_tanggal, key=lambda c: datetime.strptime(c, FORMAT_TANGGAL)
    )

    kolom_depan = KEY_COLS + EXTRA_COLS + [LINK_COL, LAST_UPDATE_COL]
    kolom_urut = [k for k in kolom_depan if k in df.columns] + kolom_tanggal_sorted
    kolom_sisa = [c for c in df.columns if c not in kolom_urut]
    return df[kolom_urut + kolom_sisa]


def main() -> None:
    if not DB_PATH.exists():
        raise FileNotFoundError(f"Database tidak ditemukan di: {DB_PATH.resolve()}")

    conn = sqlite3.connect(DB_PATH)

    df = pd.read_sql(f'SELECT * FROM "{TABLE_NAME}"', conn)
    print(f"Kolom SEBELUM dirapikan ({len(df.columns)} kolom):")
    print(list(df.columns))

    df = urutkan_kolom(df)
    print(f"\nKolom SESUDAH dirapikan ({len(df.columns)} kolom):")
    print(list(df.columns))

    df.to_sql(TABLE_NAME, conn, if_exists="replace", index=False)

    cursor = conn.cursor()
    kolom_key = ", ".join(f'"{k}"' for k in KEY_COLS)
    cursor.execute(
        f'CREATE UNIQUE INDEX IF NOT EXISTS idx_key_anomali ON "{TABLE_NAME}" ({kolom_key})'
    )
    conn.commit()
    conn.close()

    print(f"\nSelesai. {len(df)} baris tersimpan ulang dengan urutan kolom yang rapi.")


if __name__ == "__main__":
    main()
