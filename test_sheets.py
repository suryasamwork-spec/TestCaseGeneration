import os
import gspread  # type: ignore
from google.oauth2.service_account import Credentials  # type: ignore
from dotenv import load_dotenv  # type: ignore

# Load environment variables
load_dotenv()

SERVICE_ACCOUNT_FILE = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
SHEET_ID = os.getenv("GOOGLE_SHEET_ID")  # Use Sheet ID (no Drive API needed)

if not SERVICE_ACCOUNT_FILE or not SHEET_ID:
    print("[ERROR] Missing GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_SHEET_ID in .env")
    exit(1)

def test_google_sheets():
    try:
        # ✅ Only Sheets API scope needed — no Drive API required
        scope = [
            "https://www.googleapis.com/auth/spreadsheets"
        ]

        # Authenticate using service account
        creds = Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE,
            scopes=scope
        )

        # Connect to Google Sheets
        client = gspread.authorize(creds)

        # ✅ Open by Sheet ID (avoids Drive API lookup by name)
        sheet = client.open_by_key(SHEET_ID).sheet1

        # Append test row
        sheet.append_row([
            "Connection Successful",
            "Gemini Integration Ready",
            "Test Passed"
        ])

        print("[SUCCESS] Google Sheets connected successfully!")

    except Exception as e:
        print("[ERROR]", e)


if __name__ == "__main__":
    test_google_sheets()