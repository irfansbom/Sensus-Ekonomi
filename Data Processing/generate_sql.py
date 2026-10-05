import openpyxl
from pathlib import Path
from collections import defaultdict

def esc(s):
    return s.strip().upper().replace("'", "''") if s else ""

CHUNK = 35
OUT_COLS = ", ".join(
    f"b.{c}"
    for c in [
        "assignment_id", "level_2_full_code", "nama_usaha_dicari",
        "nama_usaha", "nama_usaha_edit", "assignment_status_alias",
        "keberadaan_usaha_label", "nama_usaha_sim", "nama_usaha_edit_sim",
    ]
)

def build_sql(names, kode_kab):
    blocks = []
    for n in names:
        lit = f"'{esc(n)}'"
        blocks.append(
            f"(SELECT assignment_id, level_2_full_code, {lit} AS nama_usaha_dicari, "
            "nama_usaha, nama_usaha_edit, assignment_status_alias, keberadaan_usaha_label, "
            "nama_usaha_sim, nama_usaha_edit_sim "
            "FROM (SELECT assignment_id, level_2_full_code, nama_usaha, nama_usaha_edit, "
            "assignment_status_alias, keberadaan_usaha_label, "
            f"ngram_search(nama_usaha, {lit}, 2) AS nama_usaha_sim, "
            f"ngram_search(nama_usaha_edit, {lit}, 2) AS nama_usaha_edit_sim "
            "FROM base_data) x "
            "WHERE nama_usaha_sim >= 0.8 OR nama_usaha_edit_sim >= 0.8)"
        )
    union_block = "\n    UNION ALL\n    ".join(blocks)

    return f"""WITH base_data AS (
  SELECT DISTINCT
    a.assignment_id, a.level_2_full_code,
    nama_usaha, nama_usaha_edit,
    a.assignment_status_alias, keberadaan_usaha_label
  FROM tgr_fd68e454.base_table_assignment a
  JOIN tgr_fd68e454.root_table b
    ON (a.assignment_id = b.assignment_id AND a.assignment_date_modified = b.assignment_date_modified)
  JOIN tgr_fd68e454.se2026_nested c
    ON (a.assignment_id = c.assignment_id AND a.assignment_date_modified = c.assignment_date_modified)
  WHERE a.is_active = 1
    AND a.level_2_full_code = '{kode_kab}'
)
SELECT {OUT_COLS} FROM (
  SELECT t.assignment_id, t.level_2_full_code, t.nama_usaha_dicari, t.nama_usaha, t.nama_usaha_edit, t.assignment_status_alias, t.keberadaan_usaha_label, t.nama_usaha_sim, t.nama_usaha_edit_sim, ROW_NUMBER() OVER (
    PARTITION BY t.nama_usaha_dicari
    ORDER BY GREATEST(
      t.nama_usaha_sim,
      CASE WHEN t.nama_usaha_edit_sim = '-' THEN t.nama_usaha_sim ELSE t.nama_usaha_edit_sim END
    ) DESC
  ) AS rn
  FROM (
    {union_block}
  ) t
) b
WHERE b.rn = 1;"""

wb = openpyxl.load_workbook('/mnt/user-data/uploads/Daftar_Usaha_PMDN_1671.xlsx', read_only=True, data_only=True)
ws = wb['Sheet1']
rows = ws.iter_rows(values_only=True)
header = next(rows)
idx_nama = header.index('Nama Perusahaan')
idx_kab = header.index('kode_kab')

data_per_kab = defaultdict(list)
for r in rows:
    nama = r[idx_nama]
    kab = r[idx_kab]
    if nama and str(nama).strip() and kab:
        data_per_kab[str(kab).strip()].append(str(nama).strip())

out_dir = Path('/home/claude/sql_out')
manifest = []
for kode_kab, names in data_per_kab.items():
    n_files = (len(names) - 1) // CHUNK + 1
    for i in range(0, len(names), CHUNK):
        chunk = names[i:i+CHUNK]
        fname = f"{kode_kab}_{i//CHUNK + 1}.sql"
        (out_dir / fname).write_text(build_sql(chunk, kode_kab), encoding='utf-8')
    manifest.append(f"{kode_kab}: {len(names)} rows -> {kode_kab}_1.sql ... {kode_kab}_{n_files}.sql")
    print(manifest[-1])

print("Total files:", len(list(out_dir.glob('*.sql'))))
