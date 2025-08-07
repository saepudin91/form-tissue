import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Tes Koneksi Sheets")

st.title("🔗 Tes Koneksi Google Sheets")

# Scope Google API
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

# Ambil kredensial dari secrets
try:
    creds_dict = st.secrets["gcp_service_account"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)

    # Tampilkan list spreadsheet yg bisa diakses
    st.subheader("✅ Koneksi Berhasil!")
    st.write("Berikut adalah spreadsheet yang dapat diakses:")
    
    sheets = client.openall()
    for sheet in sheets:
        st.markdown(f"- {sheet.title}")

except Exception as e:
    st.error(f"❌ Gagal konek ke Google Sheets: {e}")
