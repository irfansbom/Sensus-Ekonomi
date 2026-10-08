SELECT
    CONCAT(
        'https://fasih-sm.bps.go.id/app/assignment-detail/',
        r.assignment_id
    ) AS LINK_FASIH,
    r.assignment_id,
    r.idsbr,
    r.nama_usaha AS nama,
    r.alamat_usaha AS alamat,
    b.level_1_code AS kd_prov,
    b.level_2_code AS kd_kab,
    b.level_3_code AS kd_kec,
    b.level_4_code AS kd_desa,
    b.level_1_name AS nmprov,
    b.level_2_name AS nmkab,
    b.level_3_name AS nmkec,
    b.level_4_name AS nmdesa,
    r.keberadaan_value AS kode_flag_keberadaan,
    r.keberadaan_label AS flag_keberadaan,
	r.skala_usaha AS skala_usaha
FROM tcz_37526b20.root_table r
INNER JOIN tcz_37526b20.base_table_assignment b
    ON r.assignment_id = b.assignment_id
    AND r.assignment_date_modified = b.assignment_date_modified
WHERE b.is_active = 1 AND b.level_2_code = '10'
ORDER BY b.level_6_full_code, b.code_identity
