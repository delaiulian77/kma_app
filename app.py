import streamlit as st
from datetime import datetime
from io import BytesIO

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload


def _gcp_creds():
    # Build credentials from Streamlit secrets
    gcp_info = dict(st.secrets["gcp"])  # has type, project_id, private_key, client_email, etc.
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    return Credentials.from_service_account_info(gcp_info, scopes=scopes)


def test_write_to_sheet():
    """Append a row into the 'Logins' sheet of your KMA_DB spreadsheet."""
    creds = _gcp_creds()
    gc = gspread.authorize(creds)

    sheet_id = st.secrets["app"]["spreadsheet_id"]
    sh = gc.open_by_key(sheet_id)

    # Change this to any tab name you want to test
    ws = sh.worksheet("Logins")  # must exist

    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    row = [ts, "diag-user", "DIAG", "Spormål/Geismar/RCA-D-1435/123456789", ts]
    ws.append_row(row, value_input_option="USER_ENTERED")

    return "Row appended to 'Logins' successfully."


def test_upload_to_drive():
    """Upload a small text file to your Drive 'reports' folder."""
    creds = _gcp_creds()
    drive = build("drive", "v3", credentials=creds)

    folder_id = st.secrets["app"]["drive_folder_id"].strip()

    # 1) Sanity check: can the service account see the folder?
    try:
        folder_meta = drive.files().get(
            fileId=folder_id,
            fields="id, name, mimeType, owners"
        ).execute()
    except Exception as e:
        raise RuntimeError(
            f"Service account cannot access folder '{folder_id}'. "
            f"Check sharing and the ID. Underlying error: {e}"
        )

    # 2) Prepare content
    content = f"Hello from KMA diagnostics!\nTimestamp: {datetime.now()}\n"
    bio = BytesIO(content.encode("utf-8"))
    media = MediaIoBaseUpload(bio, mimetype="text/plain", resumable=False)

    file_metadata = {
        "name": f"diag_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
        "parents": [folder_id],
    }

    # 3) Create the file *in the folder*
    created = drive.files().create(
        body=file_metadata,
        media_body=media,
        fields="id, webViewLink"
        # do NOT set supportsAllDrives here for a personal My Drive folder
    ).execute()

    return created.get("webViewLink", "(no link returned)")


with st.expander("Diagnostics — Google Sheets & Drive", expanded=False):
    c1, c2 = st.columns(2)
    if c1.button("Test write to Sheet"):
        try:
            msg = test_write_to_sheet()
            st.success(msg)
        except Exception as e:
            st.error(f"Sheet test failed: {e}")

    if c2.button("Test upload to Drive"):
        try:
            link = test_upload_to_drive()
            st.success("Drive upload OK")
            st.write(f"Open file: {link}")
        except Exception as e:
            st.error(f"Drive test failed: {e}")
