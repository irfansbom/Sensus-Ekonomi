import streamlit as st
import sqlite3
import pandas as pd

import sqlite3

# conn = sqlite3.connect("../SQLLITE/rekap_progress_pendataan.db")
# cursor = conn.cursor()
# cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
# tables = cursor.fetchall()
# print(tables)
# conn.close()


st.title("Data Viewer")

conn = sqlite3.connect("../SQLLITE/rekap_progress_pendataan.db")
df = pd.read_sql_query("SELECT * FROM rekap_progress_pendataan", conn)
conn.close()

st.dataframe(df)
