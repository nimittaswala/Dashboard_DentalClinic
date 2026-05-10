"""
Dental AR Dashboard — Google Sheet Setup
Run this ONCE to create all tabs and headers.
Place credentials.json in C:\dental_automation\ar\ before running.
"""

import sys, os
sys.path.insert(0, r"C:\dental_automation\config")

import gspread
from google.oauth2.service_account import Credentials

try:
    from office_config import FOLDERS, AR_SETTINGS
    CREDS_PATH  = os.path.join(FOLDERS["base"], "ar", "credentials.json")
    SHEET_ID    = AR_SETTINGS["sheet_id"]
except Exception as e:
    print(f"ERROR loading config: {e}")
    sys.exit(1)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# ── Column headers for Claims Detail tab ──────────────────────────────────────
CLAIMS_HEADERS = [
    "Push Date", "Push Time", "Office",
    "Claim Num", "Patient Name", "Patient DOB", "Date of Service", "DOS To", "Date Sent",
    "Carrier", "Payor ID", "Plan Type", "Provider",
    "Procedure Codes", "Billed", "Insurance Paid", "Ins Estimate",
    "Write Off", "Claim Status", "Age (Days)", "Resubmit Count",
    "Claim ID"
]

# ── Column headers for Summary tab ────────────────────────────────────────────
SUMMARY_HEADERS = [
    "Date", "Office",
    "0-30 Days", "31-60 Days", "61-90 Days", "90+ Days",
    "Total AR", "Total Claims",
    "Oldest Claim (Days)", "Oldest Claim Date",
    "Medicaid AR", "PPO AR",
    "Claims Paid Since Last Push", "New Claims Since Last Push"
]

# ── Column headers for Notes tab ──────────────────────────────────────────────
NOTES_HEADERS = [
    "Timestamp", "Claim ID", "Office", "Patient", "Carrier",
    "Note", "Category", "Added By", "Sent To Manager", "Sent At"
]

# ── Header formatting ──────────────────────────────────────────────────────────
def format_header_row(ws):
    ws.format("1:1", {
        "backgroundColor": {"red": 0.12, "green": 0.23, "blue": 0.55},
        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
        "horizontalAlignment": "CENTER"
    })

def setup_sheet():
    print("Connecting to Google Sheets...")
    creds = Credentials.from_service_account_file(CREDS_PATH, scopes=SCOPES)
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(SHEET_ID)
    print(f"Connected to: {sh.title}")

    existing = [ws.title for ws in sh.worksheets()]
    print(f"Existing tabs: {existing}")

    # ── Overview tab ──────────────────────────────────────────────────────────
    if "Overview" not in existing:
        ws = sh.add_worksheet("Overview", rows=100, cols=20)
        print("Created: Overview")
    else:
        ws = sh.worksheet("Overview")
    ws.clear()
    ws.update("A1", [["🦷 Dental AR Dashboard — Central Overview"]])
    ws.update("A3", [["Office", "Last Updated", "0-30", "31-60", "61-90", "90+", "Total AR", "Claims", "Oldest"]])
    ws.format("A1", {"textFormat": {"bold": True, "fontSize": 14}})
    ws.format("A3:I3", {
        "backgroundColor": {"red": 0.12, "green": 0.23, "blue": 0.55},
        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}}
    })
    print("Set up: Overview")

    # ── Summary tab ───────────────────────────────────────────────────────────
    if "Summary" not in existing:
        ws = sh.add_worksheet("Summary", rows=5000, cols=len(SUMMARY_HEADERS))
        print("Created: Summary")
    else:
        ws = sh.worksheet("Summary")
    ws.clear()
    ws.append_row(SUMMARY_HEADERS)
    format_header_row(ws)
    ws.freeze(rows=1)
    print("Set up: Summary")

    # ── Claims tab ────────────────────────────────────────────────────────────
    if "Claims" not in existing:
        ws = sh.add_worksheet("Claims", rows=50000, cols=len(CLAIMS_HEADERS))
        print("Created: Claims")
    else:
        ws = sh.worksheet("Claims")
    ws.clear()
    ws.append_row(CLAIMS_HEADERS)
    format_header_row(ws)
    ws.freeze(rows=1)
    print("Set up: Claims")

    # ── Notes tab ─────────────────────────────────────────────────────────────
    if "Notes" not in existing:
        ws = sh.add_worksheet("Notes", rows=10000, cols=len(NOTES_HEADERS))
        print("Created: Notes")
    else:
        ws = sh.worksheet("Notes")
    ws.clear()
    ws.append_row(NOTES_HEADERS)
    format_header_row(ws)
    ws.freeze(rows=1)
    print("Set up: Notes")

    # ── Delete default Sheet1 if it exists ────────────────────────────────────
    if "Sheet1" in [ws.title for ws in sh.worksheets()]:
        sh.del_worksheet(sh.worksheet("Sheet1"))
        print("Removed: Sheet1")

    print("\nSheet setup complete!")
    print(f"Open your sheet at: https://docs.google.com/spreadsheets/d/{SHEET_ID}")

if __name__ == "__main__":
    setup_sheet()
