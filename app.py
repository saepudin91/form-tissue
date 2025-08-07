import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

sheet = client.open_by_key("1NGfDRnXa4rmD5n-F-ZMdtSNX__bpHiUPzKJU2KeUSaU").worksheet("Sheet1")
data = sheet.get_all_values()
print(data)
