import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
import gspread
from google.oauth2.service_account import Credentials

# =============================
# 🔑 Setup koneksi ke Google Sheets
# =============================
scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
creds = Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=scope
)
client = gspread.authorize(creds)
sheet = client.open("Log Tissue").sheet1

# ✅ Tambahkan header jika sheet kosong
if len(sheet.get_all_values()) == 0:
    header = ["Jenis", "Tanggal", "Hari", "Shift", "Pengeluaran", "Pemasukan"]
    sheet.append_row(header)

# =============================
# 📝 Konfigurasi Aplikasi
# =============================
st.set_page_config(page_title="Form Tissue", layout="wide")
st.title("📝 Form Tissue")

if "data" not in st.session_state:
    st.session_state.data = []

# =============================
# 📥 Form Input
# =============================
with st.form("form_input_shift"):
    st.subheader("Input Data Tissue")
    col1, col2 = st.columns(2)

    with col1:
        jenis = st.selectbox("Jenis Tissue", [
            "Pengeluaran Tissue Roll",
            "Pengeluaran Hand Towel",
            "Pemasukan Tissue Roll",
            "Pemasukan Hand Towel"
        ])
        tanggal = st.date_input("Tanggal", value=datetime.today())
        shift = st.selectbox("Pilih Shift", ["Shift 1", "Shift 2"])

    with col2:
        jumlah_input = st.text_input("Jumlah (contoh: 1 pcs / 2 dus / 3 roll)")

    submitted = st.form_submit_button("➕ Tambahkan")

    if submitted:
        jumlah_angka = ''.join(filter(str.isdigit, jumlah_input))
        jumlah_angka = int(jumlah_angka) if jumlah_angka else 0

        pengeluaran = jumlah_angka if "Pengeluaran" in jenis else 0
        pemasukan = jumlah_angka if "Pemasukan" in jenis else 0
        jenis_bersih = jenis.replace("Pengeluaran ", "").replace("Pemasukan ", "")

        data_baris = [jenis_bersih, tanggal.strftime("%Y-%m-%d"), tanggal.strftime("%A"), shift, pengeluaran, pemasukan]
        st.session_state.data.append(data_baris)

        try:
            sheet.append_row(data_baris)
            st.success(f"✅ Data {shift} berhasil ditambahkan & disimpan ke Google Sheets!")
        except Exception as e:
            st.error(f"❌ Gagal simpan ke Google Sheets: {e}")

# =============================
# 📊 Tampilkan Data Sheet
# =============================
try:
    records = sheet.get_all_records()
    if not records:
        st.info("Belum ada data pada Google Sheets.")
        st.stop()

    df = pd.DataFrame(records)

    st.write("📋 Data Tissue Masuk & Keluar:")
    st.dataframe(df, use_container_width=True)

    df["Tanggal"] = pd.to_datetime(df["Tanggal"])

    pengeluaran_last7 = df[(df["Pengeluaran"] > 0) & (df["Tanggal"] >= datetime.today() - timedelta(days=6))]
    pemasukan_last7 = df[(df["Pemasukan"] > 0) & (df["Tanggal"] >= datetime.today() - timedelta(days=6))]

    pengeluaran_summary = pengeluaran_last7.groupby("Jenis")["Pengeluaran"].sum().reset_index()
    pengeluaran_summary.rename(columns={"Pengeluaran": "Total Pengeluaran"}, inplace=True)

    pemasukan_summary = pemasukan_last7.groupby("Jenis")["Pemasukan"].sum().reset_index()
    pemasukan_summary.rename(columns={"Pemasukan": "Total Pemasukan"}, inplace=True)

    # 🚨 Notifikasi stok rendah
    if not pemasukan_summary.empty and not pengeluaran_summary.empty:
        st.markdown("### 🚨 Notifikasi Stok Tissue")
        stok_df = pd.merge(pemasukan_summary, pengeluaran_summary, on="Jenis", how="outer").fillna(0)
        stok_df["Sisa Stok"] = stok_df["Total Pemasukan"] - stok_df["Total Pengeluaran"]

        for _, row in stok_df.iterrows():
            if row["Sisa Stok"] <= 15:
                st.warning(f"⚠️ Stok {row['Jenis']} tersisa {int(row['Sisa Stok'])}. Segera lakukan pemesanan!")

    # =============================
    # 📥 Export ke Excel
    # =============================
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Log Tissue", index=False, startrow=2)
        worksheet = writer.sheets["Log Tissue"]

        worksheet.merge_cells("A1:F1")
        cell = worksheet["A1"]
        cell.value = "📋 Data Tissue Masuk & Keluar"
        cell.font = Font(bold=True, size=20)
        cell.alignment = Alignment(horizontal="center")

        for col in range(1, 7):
            worksheet.column_dimensions[get_column_letter(col)].width = 15

        start_rekap_row = len(df) + 6

        # 🔻 Pengeluaran
        if not pengeluaran_summary.empty:
            worksheet.cell(row=start_rekap_row, column=1).value = "🔻 Rekap Pengeluaran 7 Hari Terakhir"
            worksheet.cell(row=start_rekap_row, column=1).font = Font(bold=True, size=14)
            for r in dataframe_to_rows(pengeluaran_summary, index=False, header=True):
                worksheet.append(r)

        # 🔺 Pemasukan
        if not pemasukan_summary.empty:
            col_offset = 5
            worksheet.cell(row=start_rekap_row, column=col_offset).value = "🔺 Rekap Pemasukan 7 Hari Terakhir"
            worksheet.cell(row=start_rekap_row, column=col_offset).font = Font(bold=True, size=14)

            for idx, row in pemasukan_summary.iterrows():
                worksheet.cell(row=start_rekap_row + 1 + idx, column=col_offset).value = row["Jenis"]
                worksheet.cell(row=start_rekap_row + 1 + idx, column=col_offset + 1).value = row["Total Pemasukan"]

    buffer.seek(0)
    st.download_button(
        "📥 Download Excel Rapi",
        buffer.getvalue(),
        file_name="log_tissue_dan_rekap.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # =============================
    # 📈 Rekap di Streamlit
    # =============================
    st.markdown("---")
    st.subheader("📈 Rekap 7 Hari Terakhir")
    if not pengeluaran_summary.empty:
        st.write("### 🔻 Total Pengeluaran:")
        st.dataframe(pengeluaran_summary)
    if not pemasukan_summary.empty:
        st.write("### 🔺 Total Pemasukan:")
        st.dataframe(pemasukan_summary)

except Exception as e:
    st.error(f"Gagal ambil data dari Google Sheets: {e}")
