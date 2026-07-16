from playwright.sync_api import sync_playwright
import pandas as pd
import os
import time
from datetime import datetime

all_df = []

sql = """
    WITH data AS (
        SELECT
            DENSE_RANK() OVER (
                ORDER BY root.level_6_full_code, root.no_keluarga
            ) AS rank_keluarga,

            root.assignment_id,
            root.level_6_full_code,
            root.no_keluarga,
            root.no_bang,
            root.no_kk,
            dtsen.no_urut_kk,
            root.nama_kk,
            dtsen.nama_dtsen,
            dtsen.nik_dtsen

        FROM tgr_fd68e454.nested_dtsen dtsen
        JOIN tgr_fd68e454.root_table root
            ON root.assignment_id = dtsen.assignment_id
        WHERE dtsen.nik_dtsen IS NULL
        OR dtsen.nik_dtsen LIKE '%999999%'
    )

    SELECT *
    FROM data
    WHERE rank_keluarga BETWEEN 1 AND 10000
    ORDER BY rank_keluarga, no_urut_kk;
    """
url = ("https://fasih-dashboard.bps.go.id/api/v1/sqllab/execute/")

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    context = browser.new_context(storage_state="scrap_sqllab/session_sqllab.json")
    cookies = context.cookies()
    page = context.new_page()
    page.goto(
        "https://fasih-dashboard.bps.go.id/superset/sqllab/",
        wait_until="domcontentloaded",
        timeout=120000
    )
    
    csrf_token = page.locator("#csrf_token").get_attribute("value")
    print("CSRF:", csrf_token)
    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "referer": "https://fasih-dashboard.bps.go.id/superset/sqllab/",
        "x-csrftoken": csrf_token,
    }
    payload = {
        # "client_id": "dc9QafgUEZI",  # sementara pakai yg dari curl
        "database_id": 21,
        "json": True,
        "runAsync": False,
        "schema": "tgr_fd68e454",
        "sql": sql,
        "sql_editor_id": "944268",
        "tab": "sumsel_Missing_nik kosong",
        "tmp_table_name": "",
        "select_as_cta": False,
        "ctas_method": "TABLE",
        "queryLimit": 100000,
        "expand_data": True,
    }

    try:
        response = context.request.post(
            url,
            headers=headers,
            data=payload,
            timeout=7200000  # 2 jam
        )
        print("Status:", response.status)
        print("Content-Type:", response.headers.get("content-type"))
        print("response:", response.text())
        data = response.json()
        df = pd.DataFrame(data)
        print(f"  -> {len(df)} baris")
        folder = "./SQL Lab"
        os.makedirs(folder, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_excel = os.path.join(folder, f"missing_value_nik_sumsel_{timestamp}.xlsx")
        df.to_excel(file_excel, index=False)
        print("Total baris:", len(df))
        print("Selesai")
        print(f"File tersimpan: {file_excel}")
    except Exception as e:
        print(f"ERROR : {e}")
    time.sleep(2)
    context.close()
    browser.close()
