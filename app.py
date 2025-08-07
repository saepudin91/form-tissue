import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

# ===================== KONFIGURASI =====================
st.set_page_config(page_title="Log Tissue", layout="centered")
st.title("🧻 Form Tissue Masuk & Keluar")

# Scope dan autentikasi Google Sheets
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

try:
    creds_dict = st.secrets["gcp_service_account"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
except Exception as e:
    st.error(f"❌ Gagal mengautentikasi ke Google Sheets: {e}")
    st.stop()

# ===================== AKSES SHEET =====================
SHEET_KEY = "1NGfDRnXa4rmD5n-F-ZMdtSNX__bpHiUPzKJU2KeUSaU"
SHEET_NAME = "Sheet1"

try:
    sheet = client.open_by_key(SHEET_KEY).worksheet(SHEET_NAME)
except Exception as e:
    st.error(f"❌ Gagal membuka sheet: {e}")
    st.stop()

# ===================== CEK HEADER =====================
EXPECTED_HEADER = ["Jenis", "Tanggal", "Hari", "Shift", "Pengeluaran", "Pemasukan"]
current_data = sheet.get_all_values()

if len(current_data) == 0 or current_data[0] != EXPECTED_HEADER:
    sheet.clear()
    sheet.append_row(EXPECTED_HEADER)

# ===================== FORM =====================
with st.form("tissue_form"):
    col1, col2 = st.columns(2)

    with col1:
        jenis = st.selectbox("Jenis Tissue:", ["Tissue Roll", "Hand Towel", "Lainnya"])
        shift = st.selectbox("Shift:", ["Shift 1", "Shift 2", "Shift 3"])

    with col2:
        pengeluaran = st.number_input("Pengeluaran", min_value=0, value=0)
        pemasukan = st.number_input("Pemasukan", min_value=0, value=0)

    submitted = st.form_submit_button("➕ Tambahkan")

    if submitted:
        if pengeluaran == 0 and pemasukan == 0:
            st.warning("❗ Masukkan setidaknya pemasukan atau pengeluaran.")
        else:
            tanggal = datetime.now().strftime('%Y-%m-%d')
            hari = datetime.now().strftime('%A')

            try:
                sheet.append_row([jenis, tanggal, hari, shift, pengeluaran, pemasukan])
                st.success("✅ Data berhasil disimpan ke Google Sheets!")
            except Exception as e:
                st.error(f"❌ Gagal menyimpan data: {e}")

# ===================== TAMPILKAN DATA =====================
st.subheader("📊 Data Tissue Masuk & Keluar:")

try:
    records = sheet.get_all_values()
    if len(records) > 1:
        df = pd.DataFrame(records[1:], columns=records[0])
        st.dataframe(df, use_container_width=True)
    else:
        st.info("📭 Sheet masih kosong.")
except Exception as e:
    st.warning(f"⚠️ Gagal mengambil data: {e}")
