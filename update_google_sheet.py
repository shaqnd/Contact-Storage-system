"""
Updates the Denver Contact Directory Google Sheet with new rows from the Excel file.

Setup:
  pip install openpyxl google-auth google-auth-oauthlib google-api-python-client

Authentication (choose one):

  A) Service Account (recommended for automation):
     1. Go to console.cloud.google.com → IAM & Admin → Service Accounts
     2. Create a service account, download the JSON key
     3. Share the Google Sheet with the service account email (Editor access)
     4. Set: GOOGLE_CREDENTIALS_FILE = "path/to/service-account-key.json"
     5. Set: USE_SERVICE_ACCOUNT = True

  B) OAuth (your personal Google account):
     1. Go to console.cloud.google.com → APIs & Services → Credentials
     2. Create OAuth 2.0 Client ID (Desktop app), download the JSON
     3. Set: GOOGLE_CREDENTIALS_FILE = "path/to/oauth-client.json"
     4. Set: USE_SERVICE_ACCOUNT = False
     5. First run will open a browser to authorize; token saved to token.json
"""

import os
import openpyxl
from googleapiclient.discovery import build
from google.oauth2 import service_account
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle

# ── CONFIG ──────────────────────────────────────────────────────────────────
EXCEL_FILE = "6b2b3150-naked_denver_contact_directory_with_article_narratives.xlsx"
SPREADSHEET_ID = "1dXo4MuDB-Y2_c-TevBlN3WraffN5lc8-XYDY7Gb0DuM"
GOOGLE_CREDENTIALS_FILE = "credentials.json"  # path to your credentials JSON
USE_SERVICE_ACCOUNT = True                     # True = service account, False = OAuth
# ────────────────────────────────────────────────────────────────────────────

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Maps Excel "New_*" sheet → target Google Sheet tab name
SHEET_MAP = {
    "New_Project_Rows": "Projects",
    "New_Directory_Rows": "Directory",
    "New_Company_Rows": "Companies",
    "New_Concept_Project_Rows": "Projects",
    "New_Concept_Directory_Rows": "Directory",
    "New_Concept_Company_Rows": "Companies",
    "New_Article_Project_Rows": "Projects",
    "New_Article_Directory_Rows": "Directory",
    "New_Article_Company_Rows": "Companies",
}


def get_credentials():
    if USE_SERVICE_ACCOUNT:
        return service_account.Credentials.from_service_account_file(
            GOOGLE_CREDENTIALS_FILE, scopes=SCOPES
        )
    # OAuth flow
    creds = None
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as f:
            creds = pickle.load(f)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(GOOGLE_CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.pickle", "wb") as f:
            pickle.dump(creds, f)
    return creds


def get_existing_rows(service, sheet_name):
    """Return all values currently in the sheet (to detect duplicates)."""
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!A:A",
    ).execute()
    return [row[0] for row in result.get("values", []) if row]


def append_rows(service, sheet_name, rows):
    if not rows:
        return 0
    body = {"values": rows}
    result = service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body=body,
    ).execute()
    return result.get("updates", {}).get("updatedRows", 0)


def load_new_rows_from_excel(excel_path):
    """
    Returns dict: {target_sheet: [rows_as_lists]}
    Skips the header row of each New_* sheet.
    """
    wb = openpyxl.load_workbook(excel_path)
    collected = {}  # target_sheet → list of row lists

    for src_sheet, target_sheet in SHEET_MAP.items():
        if src_sheet not in wb.sheetnames:
            print(f"  Warning: {src_sheet} not found in Excel, skipping.")
            continue
        ws = wb[src_sheet]
        rows = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 0:
                continue  # skip header
            # Convert None → "" and all values to str
            rows.append([str(v) if v is not None else "" for v in row])
        if rows:
            collected.setdefault(target_sheet, []).extend(rows)

    return collected


def main():
    print("Loading Excel data...")
    new_data = load_new_rows_from_excel(EXCEL_FILE)
    for sheet, rows in new_data.items():
        print(f"  {sheet}: {len(rows)} new rows to append")

    print("\nConnecting to Google Sheets...")
    creds = get_credentials()
    service = build("sheets", "v4", credentials=creds)

    total_appended = 0
    for sheet_name, rows in new_data.items():
        print(f"\nAppending {len(rows)} rows to '{sheet_name}'...")

        # Deduplicate: skip rows where column A already exists in the sheet
        existing = set(get_existing_rows(service, sheet_name))
        new_rows = [r for r in rows if r[0] not in existing]
        dupes = len(rows) - len(new_rows)
        if dupes:
            print(f"  Skipped {dupes} duplicate rows (column A already present).")

        if new_rows:
            appended = append_rows(service, sheet_name, new_rows)
            print(f"  Appended {appended} rows.")
            total_appended += appended
        else:
            print("  Nothing new to add.")

    print(f"\nDone. Total rows appended: {total_appended}")


if __name__ == "__main__":
    main()
