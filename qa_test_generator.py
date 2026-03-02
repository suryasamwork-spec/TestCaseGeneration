"""
QA Test Case Generator 
======================
Workflow input -> Groq AI generates structured test cases -> Saved to Google Sheets.

Required Environment Variables (set in .env):
    GROQ_API_KEY                - Your free Groq API key
    GROQ_MODEL                  - Groq model (default: llama-3.3-70b-versatile)
    GOOGLE_SERVICE_ACCOUNT_JSON - Path to the service account JSON file
    GOOGLE_SHEET_ID             - Google Sheet ID (from the URL between /d/ and /edit)

Usage:
    python qa_test_generator.py
"""

import os
import sys
import re
import time
from datetime import date

# Load .env file automatically if present
try:
    from dotenv import load_dotenv  # type: ignore
    load_dotenv()
except ImportError:
    pass

from groq import Groq  # type: ignore
import gspread  # type: ignore
from google.oauth2.service_account import Credentials  # type: ignore


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Sheets-only scope — no Google Drive API required
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Header row to be enforced at the top of the sheet (12 columns)
HEADER_ROW = [
    "Test Case ID",       # A
    "Test Scenario",      # B
    "Test Steps",         # C
    "Expected Result",    # D
    "Actual Result",      # E  — left blank for testers to fill
    "Test Case Type",     # F  — AI-generated
    "Priority",           # G  — AI-generated
    "Status",             # H  — Pass / Fail / NS dropdown
    "Date",               # I  — Date of generation
    "Completed Date",     # J  — left blank for tester
    "Tested By",          # K  — left blank for tester
    "Comments",           # L  — left blank for tester
]

# Status dropdown options
STATUS_OPTIONS = ["Pass", "Fail", "NS"]

# Status column index (0-based) — Column H
STATUS_COL_INDEX = 7

# Groq model — loaded from .env, defaults to llama-3.3-70b-versatile
GROQ_MODEL = os.getenv("GROQ_MODEL", "llama-3.3-70b-versatile")


# ---------------------------------------------------------------------------
# 0. test_groq_connection() — Quick API sanity check
# ---------------------------------------------------------------------------

def test_groq_connection() -> bool:
    """
    Send a minimal test prompt to Groq to verify the API key and model
    are valid before running the full workflow.

    Returns:
        bool: True if connection succeeds, False otherwise.
    """
    api_key = os.getenv("GROQ_API_KEY")
    if not api_key:
        print("[Groq] ERROR: GROQ_API_KEY is not set in .env")
        return False

    model = os.getenv("GROQ_MODEL", "llama-3.3-70b-versatile")
    print(f"[Groq] Testing connection with model: {model}")

    try:
        client = Groq(api_key=api_key)
        response = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": "Say OK"}],
            max_tokens=5,
        )
        reply = response.choices[0].message.content.strip()
        print(f"[Groq] Connection OK. Model response: '{reply[:30]}'")
        return True
    except Exception as e:
        print(f"[Groq] Connection FAILED: {e}")
        return False


# ---------------------------------------------------------------------------
# 1. connect_google_sheets()
# ---------------------------------------------------------------------------

def connect_google_sheets():
    """
    Authenticate with the Google Sheets API using a service account and
    return the target worksheet object.

    Returns:
        gspread.Worksheet: The first sheet of the target spreadsheet.

    Raises:
        EnvironmentError: If required environment variables are missing.
        FileNotFoundError: If the service account JSON file does not exist.
        ConnectionError: If the sheet cannot be opened.
    """
    service_account_path = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    sheet_id = os.environ.get("GOOGLE_SHEET_ID")

    if not service_account_path:
        raise EnvironmentError(
            "Missing env var: GOOGLE_SERVICE_ACCOUNT_JSON\n"
            "Set it to the path of your service account JSON file."
        )
    if not sheet_id:
        raise EnvironmentError(
            "Missing env var: GOOGLE_SHEET_ID\n"
            "Find it in your Google Sheet URL between /d/ and /edit"
        )
    if not os.path.isfile(service_account_path):
        raise FileNotFoundError(
            f"Service account file not found: '{service_account_path}'"
        )

    print(f"[Sheets] Authenticating with: {service_account_path}")

    # Sheets-only scope — no Google Drive API required
    creds = Credentials.from_service_account_file(service_account_path, scopes=SCOPES)
    client = gspread.authorize(creds)

    # Open by Sheet ID — no Drive API lookup needed
    try:
        spreadsheet = client.open_by_key(sheet_id)
    except Exception as e:
        raise ConnectionError(
            f"Could not open sheet ID '{sheet_id}'.\n"
            "Check:\n"
            "  1. Sheet ID is correct (from URL between /d/ and /edit)\n"
            "  2. Service account email has Editor access to the sheet\n"
            "  3. Google Sheets API is enabled in your Cloud project\n"
            f"Details: {e}"
        )

    worksheet = spreadsheet.sheet1
    print(f"[Sheets] Connected to sheet ID: {sheet_id}")
    return worksheet


# ---------------------------------------------------------------------------
# 2. generate_test_cases_with_ai(workflow)
# ---------------------------------------------------------------------------

def generate_test_cases_with_ai(workflow: str) -> list:
    """
    Send the workflow description to Groq AI and parse structured QA test cases.

    Args:
        workflow (str): A plain-text workflow description.

    Returns:
        list[list[str]]: Each row: [TC_ID, Name, Steps, Expected Result, Priority]

    Raises:
        EnvironmentError: If GROQ_API_KEY is not set.
        ValueError: If the response cannot be parsed into test cases.
        Exception: Propagates API errors with a descriptive message.
    """
    api_key = os.environ.get("GROQ_API_KEY")
    if not api_key:
        raise EnvironmentError(
            "Missing env var: GROQ_API_KEY\n"
            "Get a free key from https://console.groq.com/keys"
        )

    # Initialize Groq client
    client = Groq(api_key=api_key)

    system_prompt = (
        "You are a Senior QA Automation Engineer. "
        "You generate comprehensive, exhaustive, multi-level QA test cases strictly in pipe-separated format. "
        "No markdown, no headers, no explanations — only test case lines."
    )

    # Prompt: AI outputs 6 fields—TC_Type is generated by AI
    user_prompt = f"""You are a Senior QA Automation Engineer.

Generate comprehensive and detailed QA test cases for the following workflow.

WORKFLOW:
{workflow}

IMPORTANT REQUIREMENTS:
- The workflow contains multiple modules. Each module has subfolders with UI elements.
- You MUST generate test cases at ALL of these levels:
  LEVEL 1: Overall dashboard (load, layout, navigation, sidebar, responsive design)
  LEVEL 2: Every module (navigation, module load, permissions)
  LEVEL 3: Every subfolder/section inside each module (e.g. Asset List, Add Asset, Asset Details)
  LEVEL 4: Every UI component inside each subfolder (tables, forms, buttons, dropdowns, filters, charts)

- For EACH subfolder, generate tests covering:
  * Navigation and page load
  * UI elements visibility and layout
  * Each button (submit, cancel, edit, delete, export, search, filter, sort)
  * Form field validation (required, min/max length, format, special chars)
  * Table operations (column sort, row click, pagination, empty state)
  * Search: valid query, empty query, special characters, partial match, no results
  * Filter: single filter, multiple filters, clear filter, invalid filter
  * Create operation: valid data, duplicate, boundary values
  * Edit operation: update valid data, cancel edit, concurrent edit
  * Delete operation: confirm delete, cancel delete, delete referenced record
  * Export: CSV/PDF download, exported data accuracy
  * Error handling: API errors, network failure, session timeout
  * Permission: admin access, non-admin restriction, read-only user
  * Edge cases: empty lists, very long strings, SQL injection attempt, XSS

- Generate at MINIMUM 15-25 test cases per major module.
- There is NO LIMIT on total test cases. Exhaustive coverage is required.
- Make test steps highly specific (mention actual field names, button labels, menu items from the workflow).

OUTPUT FORMAT RULES (STRICTLY FOLLOW):
- Each test case on its own line.
- Output EXACTLY 6 pipe "|" separated fields:
  TC_ID | Test Scenario | Test Steps | Expected Result | Test Case Type | Priority
- TC_IDs must be unique and sequential: TC001, TC002, ...
- Test Scenario: Concise, natural title. Do NOT start with "Verify".
- Test Steps: At least 3 specific, detailed actions separated by semicolons.
- Expected Result: Precise, measurable outcome.
- Test Case Type: EXACTLY one of — Functional, Negative, UI/UX, Validation
- Priority: EXACTLY one of — High, Medium, Low
- No header row. No markdown. No bullets. One test case per line only.

EXAMPLE OUTPUT (follow this depth and specificity):
TC001 | Successful admin login | Navigate to /login; Enter admin@company.com in Email field; Enter correct password; Click 'Login' button | Dashboard home page loads; Sidebar shows all admin modules; User avatar visible top-right | Functional | High
TC002 | Login with wrong password | Navigate to /login; Enter valid email; Enter incorrect password; Click 'Login' button | Error toast shows 'Invalid credentials'; Password field highlighted red; User remains on login page | Negative | High
TC003 | Asset list table pagination | Navigate to Asset Management > Asset List; Scroll to bottom of table; Click 'Next Page' button | Page 2 of assets loads; Row count shows correct range (e.g. 11-20 of 100); Previous button becomes active | Functional | Medium
TC004 | Add Asset form - name field max length | Navigate to Asset Management > Add Asset; Enter 256 characters in the Asset Name field; Click Submit | Validation error: 'Asset Name must not exceed 255 characters'; Form is not submitted | Validation | High
TC005 | Export asset list to CSV | Navigate to Asset Management > Asset List; Apply any filter; Click 'Export CSV' button | CSV file downloads with filtered records; Column headers match table; Date format is correct | Functional | Medium

Now generate exhaustive test cases for the workflow above:"""


    model = os.getenv("GROQ_MODEL", "llama-3.3-70b-versatile")
    print(f"[Groq] Sending workflow to {model}...")

    try:
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.2, # Low temp for structured adherence
        )
    except Exception as e:
        raise Exception(
            f"[Groq] API call failed: {e}\n"
            "Check your GROQ_API_KEY and network connection."
        )

    raw_text = response.choices[0].message.content.strip()

    if not raw_text:
        raise ValueError(
            "[Groq] Received an empty response. "
            "Try rephrasing your workflow description."
        )

    print("[Groq] Response received. Parsing test cases...")

    # Parse each line into an 8-element row
    test_cases = []
    for line_num, line in enumerate(raw_text.splitlines(), start=1):
        line = line.strip()

        # Skip blank lines or markdown fences
        if not line or line.startswith("```") or line.startswith("#"):
            continue

        parts = [p.strip() for p in line.split("|")]

        # Expect exactly 6 fields from the LLM:
        # TC_ID | Scenario | Steps | Expected Result | Test Case Type | Priority
        if len(parts) != 6:
            print(
                f"[Groq] Warning: Skipping line {line_num} "
                f"(expected 6 fields, got {len(parts)}): {line!r}"
            )
            continue

        # Validate TC ID format (e.g. TC001)
        if not re.match(r"^TC\d+$", parts[0], re.IGNORECASE):
            print(
                f"[Groq] Warning: Skipping line {line_num} "
                f"(invalid TC ID '{parts[0]}'): {line!r}"
            )
            continue

        # Format steps (index 2) into a clean numbered list
        raw_steps = parts[2]
        step_items = [s.strip() for s in raw_steps.split(";") if s.strip()]
        if step_items:
            parts[2] = "\n".join(f"{i}. {step}" for i, step in enumerate(step_items, 1))

        # Build the full 12-column row:
        # [TC_ID, Scenario, Steps, Expected, <<Actual blank>>, TC_Type (AI), Priority, <<Status blank>>, Date, <<Completed Date blank>>, <<Tested By blank>>, <<Comments blank>>]
        row = [
            parts[0],  # Test Case ID
            parts[1],  # Test Scenario
            parts[2],  # Test Steps (numbered)
            parts[3],  # Expected Result
            "",        # Actual Result — blank for tester
            parts[4],  # Test Case Type — AI-generated
            parts[5],  # Priority — AI-generated
            "",        # Status — blank dropdown
            date.today().strftime("%d/%m/%Y"),  # Date
            "",        # Completed Date — blank for tester
            "",        # Tested By — blank for tester
            "",        # Comments — blank for tester
        ]

        test_cases.append(row)

    if not test_cases:
        raise ValueError(
            "[Groq] Could not parse any valid test cases.\n"
            f"Raw response:\n{raw_text}"
        )

    print(f"[Groq] Parsed {len(test_cases)} test case(s) successfully.")
    return test_cases


# ---------------------------------------------------------------------------
# 3. ensure_headers(sheet)
# ---------------------------------------------------------------------------

def ensure_headers(sheet: gspread.Worksheet) -> bool:
    """
    Check if the Google Sheet has the required headers.
    If empty or missing headers, append the header row safely.
    Returns True if headers were added (new sheet), False if they already existed.
    """
    try:
        # Get the first row to check for existing headers
        first_row = sheet.row_values(1)
        
        # If sheet is not empty and first row matches our headers, do nothing
        if first_row and first_row[0].strip().lower() == HEADER_ROW[0].lower():
            print("[Sheets] Headers already exist")
            return False
            
    except Exception:
        # Sheet might be completely empty, triggering an exception on row_values(1)
        pass

    print("[Sheets] Adding headers")
    
    if sheet.row_count == 0 or not sheet.get_all_values():
        sheet.append_row(HEADER_ROW, value_input_option="RAW")
    else:
        # Insert at row 1 without overwriting existing data
        sheet.insert_row(HEADER_ROW, index=1, value_input_option="RAW")
        
    return True


# ---------------------------------------------------------------------------
# 3b. add_status_validation(sheet)
# ---------------------------------------------------------------------------

def add_status_validation(sheet: gspread.Worksheet) -> None:
    """
    Apply a Pass / Fail / NS dropdown validation to the entire Status column
    (column F, index 5), starting from row 2 (skipping the header).
    Uses the Google Sheets API batchUpdate via gspread.
    """
    spreadsheet = sheet.spreadsheet
    sheet_id = sheet.id

    # Build ONE_OF_LIST dropdown values
    dropdown_values = [{"userEnteredValue": opt} for opt in STATUS_OPTIONS]

    body = {
        "requests": [
            {
                "setDataValidation": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 1,          # Row 2 (0-based), skip header
                        "startColumnIndex": STATUS_COL_INDEX,
                        "endColumnIndex": STATUS_COL_INDEX + 1,
                    },
                    "rule": {
                        "condition": {
                            "type": "ONE_OF_LIST",
                            "values": dropdown_values,
                        },
                        "showCustomUi": True,   # Show as dropdown arrow
                        "strict": False,         # Allow typing other values
                    },
                }
            }
        ]
    }

    try:
        spreadsheet.batch_update(body)
        print("[Sheets] Status dropdown validation applied (Pass / Fail / NS)")
    except Exception as e:
        # Non-fatal — sheet works fine without dropdown
        print(f"[Sheets] Warning: Could not apply dropdown validation: {e}")


# ---------------------------------------------------------------------------
# 3c. add_test_case_type_validation(sheet)
# ---------------------------------------------------------------------------

TC_TYPE_OPTIONS = ["Functional", "Positive", "Negative", "UI/UX", "Security"]
TC_TYPE_COL_INDEX = 5

def add_test_case_type_validation(sheet: gspread.Worksheet) -> None:
    """
    Apply a dropdown validation to the Test Case Type column so testers
    can pick dynamically: Functional / Positive / Negative / UI/UX / Security.
    """
    spreadsheet = sheet.spreadsheet
    sheet_id = sheet.id

    dropdown_values = [{"userEnteredValue": opt} for opt in TC_TYPE_OPTIONS]

    body = {
        "requests": [
            {
                "setDataValidation": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 1,              # Skip header
                        "startColumnIndex": TC_TYPE_COL_INDEX,
                        "endColumnIndex":   TC_TYPE_COL_INDEX + 1,
                    },
                    "rule": {
                        "condition": {
                            "type": "ONE_OF_LIST",
                            "values": dropdown_values,
                        },
                        "showCustomUi": True,   # Show dropdown arrow
                        "strict": False,         # Allow other values too
                    },
                }
            }
        ]
    }

    try:
        spreadsheet.batch_update(body)
        print("[Sheets] Test Case Type dropdown applied (Functional/Positive/Negative/UI/UX/Security)")
    except Exception as e:
        print(f"[Sheets] Warning: Could not apply Test Case Type dropdown: {e}")


# ---------------------------------------------------------------------------
# 3d. clear_conditional_formatting  /  3e. apply_subtle_tc_type_colors
# ---------------------------------------------------------------------------

def clear_conditional_formatting(sheet: gspread.Worksheet) -> None:
    """Delete ALL conditional format rules on this sheet (cleans up stale rules)."""
    spreadsheet = sheet.spreadsheet
    sheet_id    = sheet.id
    info        = spreadsheet.fetch_sheet_metadata()
    target = next(
        (s for s in info.get("sheets", []) if s["properties"]["sheetId"] == sheet_id),
        None,
    )
    if target is None:
        return
    n = len(target.get("conditionalFormats", []))
    if n == 0:
        return
    reqs = [
        {"deleteConditionalFormatRule": {"sheetId": sheet_id, "index": i}}
        for i in range(n - 1, -1, -1)
    ]
    try:
        spreadsheet.batch_update({"requests": reqs})
        print(f"[Sheets] Cleared {n} old conditional format rule(s).")
    except Exception as e:
        print(f"[Sheets] Warning: Could not clear old format rules: {e}")


# Subtle professional palette: soft pastel bg + matching dark text
_TC_PALETTE = {
    "Functional": (
        {"red": 0.98, "green": 0.94, "blue": 0.82},  # warm cream
        {"red": 0.55, "green": 0.30, "blue": 0.00},  # dark amber text
    ),
    "Negative": (
        {"red": 0.99, "green": 0.87, "blue": 0.87},  # soft rose
        {"red": 0.60, "green": 0.08, "blue": 0.08},  # dark red text
    ),
    "UI/UX": (
        {"red": 0.84, "green": 0.92, "blue": 1.00},  # sky blue
        {"red": 0.05, "green": 0.28, "blue": 0.58},  # dark blue text
    ),
    "Validation": (
        {"red": 0.86, "green": 0.97, "blue": 0.87},  # mint
        {"red": 0.06, "green": 0.40, "blue": 0.10},  # dark green text
    ),
    "Positive": (
        {"red": 0.86, "green": 0.97, "blue": 0.87},  # mint
        {"red": 0.06, "green": 0.40, "blue": 0.10},  # dark green text
    ),
    "Security": (
        {"red": 0.93, "green": 0.87, "blue": 0.99},  # soft lavender
        {"red": 0.32, "green": 0.08, "blue": 0.55},  # dark purple text
    ),
}


def apply_subtle_tc_type_colors(sheet: gspread.Worksheet) -> None:
    """Apply soft pastel conditional formatting to Test Case Type column."""
    spreadsheet = sheet.spreadsheet
    sheet_id    = sheet.id
    reqs = []
    for label, (bg, fg) in _TC_PALETTE.items():
        reqs.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{
                        "sheetId":         sheet_id,
                        "startRowIndex":   1,
                        "startColumnIndex": TC_TYPE_COL_INDEX,
                        "endColumnIndex":   TC_TYPE_COL_INDEX + 1,
                    }],
                    "booleanRule": {
                        "condition": {
                            "type":   "TEXT_EQ",
                            "values": [{"userEnteredValue": label}],
                        },
                        "format": {
                            "backgroundColor": bg,
                            "textFormat": {"bold": False, "foregroundColor": fg},
                        },
                    },
                },
                "index": 0,
            }
        })
    try:
        spreadsheet.batch_update({"requests": reqs})
        print("[Sheets] Subtle professional colors applied to Test Case Type.")
    except Exception as e:
        print(f"[Sheets] Warning: Could not apply colors: {e}")


# ---------------------------------------------------------------------------
# 4. write_test_cases_to_sheets(sheet, test_cases)
# ---------------------------------------------------------------------------

def write_test_cases_to_sheets(sheet, test_cases: list) -> None:
    """
    Append QA test cases in chunks to avoid Google Sheets 429 rate-limit.
    Writes BATCH_SIZE rows per API call, with retry + back-off on each chunk.
    """
    BATCH_SIZE = 20      # Rows per API call (safe limit)
    MAX_RETRIES = 4      # Retry attempts per chunk
    BASE_SLEEP  = 12     # Seconds to wait between successful batches

    total = len(test_cases)
    print(f"[Sheets] Writing {total} test cases in batches of {BATCH_SIZE}...")

    for start in range(0, total, BATCH_SIZE):
        chunk = test_cases[start : start + BATCH_SIZE]
        end   = min(start + BATCH_SIZE, total)

        for attempt in range(1, MAX_RETRIES + 1):
            try:
                sheet.append_rows(chunk, value_input_option="RAW")
                print(f"  [+] Wrote rows {start + 1}–{end} / {total}")
                break   # success
            except Exception as e:
                err_str = str(e)
                if "429" in err_str or "Quota" in err_str:
                    wait = BASE_SLEEP * attempt  # exponential-ish back-off
                    print(f"  [!] Rate limit hit (attempt {attempt}/{MAX_RETRIES}). "
                          f"Waiting {wait}s before retry...")
                    time.sleep(wait)
                else:
                    # Non-quota error — raise immediately
                    raise
        else:
            raise RuntimeError(
                f"Failed to write rows {start + 1}–{end} after {MAX_RETRIES} retries."
            )

        # Pause between batches to stay under quota
        if end < total:
            print(f"  [~] Pausing {BASE_SLEEP}s before next batch...")
            time.sleep(BASE_SLEEP)

    print(f"[Sheets] Done. All {total} test cases written successfully.")


# ---------------------------------------------------------------------------
# 4. main()
# ---------------------------------------------------------------------------

def main():
    """
    Main entry point:
    1. Verify Groq API connection.
    2. Get workflow from user input.
    3. Connect to Google Sheets & ensure headers.
    4. Generate test cases with Groq API.
    5. Write test cases to the sheet safely.
    """
    print("=" * 60)
    print("  QA Test Case Generator - Powered by Groq AI")
    print(f"  Model: {GROQ_MODEL}")
    print("=" * 60)

    # --- Step 0: Verify Groq API ---
    print("\n[Groq] Verifying API connection...")
    if not test_groq_connection():
        print("\n[Error] Groq API connection failed.")
        print("Check GROQ_API_KEY in your .env file.")
        print("Get a free key: https://console.groq.com/keys")
        sys.exit(1)
    print()

    # --- Step 1: Get workflow input ---
    print("Enter your workflow description below.")
    print("Press Enter twice when finished.\n")

    lines = []
    try:
        while True:
            line = input()
            if line == "" and lines and lines[-1] == "":
                break
            lines.append(line)
    except KeyboardInterrupt:
        print("\n\n[Info] Cancelled. Exiting.")
        sys.exit(0)

    workflow = "\n".join(lines).strip()

    if not workflow:
        print("\n[Error] Workflow input cannot be empty. Please run again.")
        sys.exit(1)

    print(f"\n[Info] Workflow captured ({len(workflow)} characters).\n")

    # --- Step 2: Connect to Google Sheets ---
    try:
        sheet = connect_google_sheets()
        print("[Sheets] Connected successfully")
        ensure_headers(sheet)

        # Always clear stale format rules then reapply fresh professional ones
        clear_conditional_formatting(sheet)
        add_status_validation(sheet)
        add_test_case_type_validation(sheet)
        apply_subtle_tc_type_colors(sheet)

    except (EnvironmentError, FileNotFoundError, ConnectionError) as e:
        print(f"\n[Error] {e}")
        sys.exit(1)
    except Exception as e:
        print(f"\n[Error] Google Sheets connection failed: {e}")
        sys.exit(1)

    # --- Step 3: Generate test cases with Groq ---
    try:
        test_cases = generate_test_cases_with_ai(workflow)
    except EnvironmentError as e:
        print(f"\n[Error] {e}")
        sys.exit(1)
    except ValueError as e:
        print(f"\n[Error] {e}")
        sys.exit(1)
    except Exception as e:
        print(f"\n[Error] Groq API error: {e}")
        sys.exit(1)

    # --- Step 4: Write to Google Sheets ---
    try:
        write_test_cases_to_sheets(sheet, test_cases)
    except Exception as e:
        print(f"\n[Error] Failed to write to Google Sheets: {e}")
        sys.exit(1)

    # --- Done ---
    print("\n" + "=" * 60)
    print(f"  SUCCESS! {len(test_cases)} test case(s) saved to Google Sheets.")
    print("=" * 60)


if __name__ == "__main__":
    main()
