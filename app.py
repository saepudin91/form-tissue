import streamlit as st
from oauth2client.service_account import ServiceAccountCredentials

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds_dict = st.secrets["gcp_service_account"]
credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
