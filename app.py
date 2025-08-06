import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Form Tissue", layout="wide")
st.title("📝 Form Tissue")

if "data" not in st.session_state:
    st.session_state.data = []

def write_to_gsheet(df, sheet_name="Log Tissue Online"):
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]

    creds_dict = {
        "type": st.secrets["gcp_service_account"]["type"],
        "project_id": st.secrets["gcp_service_account"]["project_id"],
        "private_key_id": st.secrets["gcp_service_account"]["private_key_id"],
        "private_key": st.secrets["gcp_service_account"]["private_key"].replace("\\n", "\n"),
        "client_email": st.secrets["gcp_service_account"]["client_email"],
        "client_id": st.secrets["gcp_service_account"]["client_id"],
        "auth_uri": st.secrets["gcp_service_account"]["auth_uri"],
        "token_uri": st.secrets["gcp_service_account"]["token_uri"],
        "auth_provider_x509_cert_url": st.secrets["gcp_service_account"]["auth_provider_x509_cert_url"],
        "client_x509_cert_url": st.secrets["gcp_service_account"]["client_x509_cert_url"],
    }

    credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(credentials)

    try:
        spreadsheet = client.open(sheet_name)
    except:
        spreadsheet = client.create(sheet_name)

    worksheet = spreadsheet.sheet1
    worksheet.clear()

    # Insert headers
    worksheet.append_row(df.columns.tolist())

    # Insert data rows
    for row in df.values.tolist():
        worksheet.append_row(row)

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

        st.session_state.data.append({
            "Jenis": jenis,
            "Tanggal": tanggal.strftime("%Y-%m-%d"),
            "Hari": tanggal.strftime("%A"),
            "Shift": shift,
            "Jumlah": jumlah_angka
        })

        df_all = pd.DataFrame(st.session_state.data)

        try:
            write_to_gsheet(df_all)
            st.success("✅ Data berhasil ditambahkan & disimpan ke Google Sheets!")
        except Exception as e:
            st.error(f"❌ Gagal simpan ke Google Sheets: {e}")

if st.session_state.data:
    df = pd.DataFrame(st.session_state.data)

    df["Pengeluaran"] = df.apply(lambda row: row["Jumlah"] if "Pengeluaran" in row["Jenis"] else 0, axis=1)
    df["Pemasukan"] = df.apply(lambda row: row["Jumlah"] if "Pemasukan" in row["Jenis"] else 0, axis=1)
    df["Jenis"] = df["Jenis"].str.replace("Pengeluaran ", "").str.replace("Pemasukan ", "")
    df_output = df.drop(columns=["Jumlah"])

    st.write("📋 Data Tissue Masuk & Keluar:")
    st.dataframe(df_output, use_container_width=True)

    df_rekap = df.copy()
    df_rekap["Tanggal"] = pd.to_datetime(df_rekap["Tanggal"])

    pengeluaran_last7 = df_rekap[(df_rekap["Pengeluaran"] > 0) & 
                                 (df_rekap["Tanggal"] >= datetime.today() - timedelta(days=6))]
    pemasukan_last7 = df_rekap[(df_rekap["Pemasukan"] > 0) & 
                               (df_rekap["Tanggal"] >= datetime.today() - timedelta(days=6))]

    pengeluaran_summary = pengeluaran_last7.groupby("Jenis")["Pengeluaran"].sum().reset_index()
    pengeluaran_summary.rename(columns={"Pengeluaran": "Total Pengeluaran"}, inplace=True)

    pemasukan_summary = pemasukan_last7.groupby("Jenis")["Pemasukan"].sum().reset_index()
    pemasukan_summary.rename(columns={"Pemasukan": "Total Pemasukan"}, inplace=True)

    if not pemasukan_summary.empty and not pengeluaran_summary.empty:
        st.markdown("### 🚨 Notifikasi Stok Tissue")
        stok_df = pd.merge(pemasukan_summary, pengeluaran_summary, on="Jenis", how="outer").fillna(0)
        stok_df["Sisa Stok"] = stok_df["Total Pemasukan"] - stok_df["Total Pengeluaran"]

        for _, row in stok_df.iterrows():
            if row["Sisa Stok"] <= 15:
                st.warning(f"⚠️ Stok {row['Jenis']} tersisa {int(row['Sisa Stok'])}. Segera lakukan pemesanan!")

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_output.to_excel(writer, sheet_name="Log Tissue", index=False, startrow=2)
        workbook = writer.book
        worksheet = writer.sheets["Log Tissue"]

        worksheet.merge_cells("A1:F1")
        cell = worksheet["A1"]
        cell.value = "📋 Data Tissue Masuk & Keluar"
        cell.font = Font(bold=True, size=20)
        cell.alignment = Alignment(horizontal="center")

        for col in range(1, 7):
            worksheet.column_dimensions[get_column_letter(col)].width = 15

        start_row = len(df_output) + 6

        if not pengeluaran_summary.empty:
            worksheet.cell(row=start_row, column=1).value = "🔻 Rekap Pengeluaran 7 Hari Terakhir"
            worksheet.cell(row=start_row, column=1).font = Font(bold=True, size=14)

            for i, r in enumerate(dataframe_to_rows(pengeluaran_summary, index=False, header=True)):
                for j, val in enumerate(r, start=1):
                    worksheet.cell(row=start_row + i + 1, column=j, value=val)

        if not pemasukan_summary.empty:
            start_col = 5
            worksheet.cell(row=start_row, column=start_col).value = "🔺 Rekap Pemasukan 7 Hari Terakhir"
            worksheet.cell(row=start_row, column=start_col).font = Font(bold=True, size=14)

            for i, r in enumerate(dataframe_to_rows(pemasukan_summary, index=False, header=True)):
                for j, val in enumerate(r, start=1):
                    worksheet.cell(row=start_row + i + 1, column=start_col + j - 1, value=val)

    buffer.seek(0)

    st.download_button(
        "📥 Download Excel Rapi",
        buffer.getvalue(),
        file_name="log_tissue_dan_rekap.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.markdown("---")
    st.subheader("📈 Rekap 7 Hari Terakhir")

    if not pengeluaran_summary.empty:
        st.write("### 🔻 Total Pengeluaran:")
        st.dataframe(pengeluaran_summary)

    if not pemasukan_summary.empty:
        st.write("### 🔺 Total Pemasukan:")
        st.dataframe(pemasukan_summary)

else:
    st.info("Belum ada data yang ditambahkan.")
