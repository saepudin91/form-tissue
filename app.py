import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

st.set_page_config(page_title="Log Tissue", layout="centered")
st.title("🧻 Form Tissue Masuk & Keluar")

# Gunakan secrets di Streamlit Cloud
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds_dict = st.secrets["gcp_service_account"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)

# Ganti dengan nama Google Sheet kamu
sheet = client.open("Log Tissue").sheet1

with st.form("tissue_form"):
    col1, col2 = st.columns(2)
    with col1:
        jenis = st.selectbox("Jenis Tissue:", ["Tissue Roll", "Hand Towel", "Lainnya"])
        shift = st.selectbox("Shift:", ["Shift 1", "Shift 2", "Shift 3"])
    with col2:
        tanggal = datetime.today().strftime('%Y-%m-%d')
        hari = datetime.today().strftime('%A')
        pengeluaran = st.number_input("Pengeluaran", min_value=0, value=0)
        pemasukan = st.number_input("Pemasukan", min_value=0, value=0)

    if st.form_submit_button("➕ Tambahkan"):
        try:
            sheet.append_row([jenis, tanggal, hari, shift, pengeluaran, pemasukan])
            st.success("✅ Data berhasil disimpan ke Google Sheets!")
        except Exception as e:
            st.error(f"❌ Gagal menyimpan data: {e}")

st.subheader("📊 Data Tissue Masuk & Keluar:")
try:
    df = pd.DataFrame(sheet.get_all_records())
    st.dataframe(df, use_container_width=True)
except Exception as e:
    st.warning(f"⚠️ Gagal mengambil data: {e}")
