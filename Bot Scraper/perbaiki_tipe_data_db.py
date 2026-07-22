"""
Script sekali jalan untuk memperbaiki kolom numerik di database SQLite yang
masih tersimpan sebagai float (muncul '.0') menjadi integer.

Jalankan ini SEKALI setelah update script utama, untuk membersihkan data lama.
Tidak perlu dijalankan lagi setelah ini kalau script utama sudah pakai
konversi Int64 sebelum upsert.
"""

import sqlite3
from pathlib import Path

DB_PATH = Path("../SQLLITE/rekap_progress_pendataan.db")
TABLE_NAME = "rekap_progress_pendataan"

# Kolom yang TIDAK perlu dikonversi (bukan angka)
KOLOM_BUKAN_ANGKA = {"id_wilayah", "kd_kab", "nama_sls", "PPL", "PML", "last_update"}

# Kolom yang mau dihapus total dari database
KOLOM_DIHAPUS = [
    "Jumlah Keluarga Prelist Awal",
    "Progres Pendataan Usaha dalam Keluarga",
    "Jumlah Keluarga Menolak Didata",
]


def hapus_kolom(kolom_dihapus: list[str]) -> None:
    """Hapus kolom dari tabel. Coba DROP COLUMN langsung dulu (SQLite >= 3.35);
    kalau versi SQLite-nya lebih lama dan tidak mendukung, otomatis fallback
    ke cara manual (buat tabel baru tanpa kolom itu, pindahkan data, ganti nama)."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute(f'PRAGMA table_info("{TABLE_NAME}")')
    kolom_ada = {row[1] for row in cursor.fetchall()}
    kolom_valid = [k for k in kolom_dihapus if k in kolom_ada]
    kolom_tidak_ada = [k for k in kolom_dihapus if k not in kolom_ada]

    if kolom_tidak_ada:
        print(f"Kolom berikut tidak ditemukan di tabel (dilewati): {kolom_tidak_ada}")

    if not kolom_valid:
        print("Tidak ada kolom valid untuk dihapus.")
        conn.close()
        return

    try:
        for kolom in kolom_valid:
            cursor.execute(f'ALTER TABLE "{TABLE_NAME}" DROP COLUMN "{kolom}"')
            print(f"  Kolom '{kolom}' berhasil dihapus (DROP COLUMN langsung).")
        conn.commit()

    except sqlite3.OperationalError as e:
        # Kemungkinan versi SQLite belum mendukung DROP COLUMN.
        print(f"DROP COLUMN langsung gagal ({e}), pakai cara manual (rebuild tabel)...")
        conn.rollback()

        kolom_dipertahankan = [k for k in kolom_ada if k not in kolom_valid]
        kolom_str = ", ".join(f'"{k}"' for k in kolom_dipertahankan)

        cursor.execute(f'ALTER TABLE "{TABLE_NAME}" RENAME TO "{TABLE_NAME}_lama"')
        cursor.execute(f'CREATE TABLE "{TABLE_NAME}" AS SELECT {kolom_str} FROM "{TABLE_NAME}_lama"')
        cursor.execute(
            f'CREATE UNIQUE INDEX IF NOT EXISTS idx_id_wilayah ON "{TABLE_NAME}" (id_wilayah)'
        )
        cursor.execute(f'DROP TABLE "{TABLE_NAME}_lama"')
        conn.commit()
        print(f"  Berhasil hapus {len(kolom_valid)} kolom lewat rebuild tabel.")

    conn.close()


def perbaiki_tipe_data():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute(f'PRAGMA table_info("{TABLE_NAME}")')
    semua_kolom = [row[1] for row in cursor.fetchall()]
    kolom_angka = [k for k in semua_kolom if k not in KOLOM_BUKAN_ANGKA]

    print(f"Kolom yang akan dibersihkan: {kolom_angka}")

    for kolom in kolom_angka:
        # CAST(... AS INTEGER) di SQLite otomatis membuang '.0', tapi tetap
        # menjaga NULL sebagai NULL (tidak dipaksa jadi 0).
        cursor.execute(f"""
            UPDATE "{TABLE_NAME}"
            SET "{kolom}" = CAST("{kolom}" AS INTEGER)
            WHERE "{kolom}" IS NOT NULL
        """)
        print(f"  Kolom '{kolom}' selesai dibersihkan ({cursor.rowcount} baris kena update).")

    conn.commit()
    conn.close()
    print("Selesai. Cek ulang datanya, seharusnya sudah tidak ada '.0' lagi.")


if __name__ == "__main__":
    hapus_kolom(KOLOM_DIHAPUS)
    perbaiki_tipe_data()