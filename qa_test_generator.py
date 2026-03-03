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

# Header row to be enforced at the top of the sheet (13 columns)
HEADER_ROW = [
    "Test Case ID",       # A
    "Module",             # B  — feature/module group
    "Test Scenario",      # C
    "Test Steps",         # D
    "Expected Result",    # E
    "Actual Result",      # F  — left blank for testers to fill
    "Test Case Type",     # G  — AI-generated
    "Priority",           # H  — AI-generated
    "Status",             # I  — Pass / Fail / NS dropdown
    "Date",               # J  — Date of generation
    "Completed Date",     # K  — left blank for tester
    "Tested By",          # L  — left blank for tester
    "Comments",           # M  — left blank for tester
]

# Status dropdown options
STATUS_OPTIONS = ["Pass", "Fail", "NS"]

# Status column index (0-based) — Column I (index 8)
STATUS_COL_INDEX = 8

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
        "You are a Senior QA Engineer with 10+ years of experience testing enterprise dashboards and web applications. "
        "You generate HIGH-QUALITY, NON-REPETITIVE, human-written QA test cases in strict pipe-separated format. "
        "No markdown, no headers, no explanations — only test case lines. "
        "MINDSET: Think like an experienced human tester — focus on real business risk, not mechanical checklists. "
        "BEFORE generating: (1) Identify all unique features and behaviors. "
        "(2) Map exactly ONE test case per unique behavior — never rephrase the same logic. "
        "(3) If two scenarios validate the same logic, keep only the more valuable one. "
        "(4) Do NOT create artificial variations. "
        "(5) Limit total output to 40–45 test cases maximum — stop once all unique behaviors are covered. "
        "CRITICAL RULE — Test Steps: Do NOT start with navigation. Steps begin at the direct action being tested. "
        "CRITICAL RULE — Expected Result: Focus on the FINAL action only. One measurable outcome. "
        "Never repeat the same expected result across different test cases."
    )

    # Prompt: AI outputs 7 fields — TC_ID | Module | Scenario | Steps | Expected Result | TC Type | Priority
    user_prompt = f"""You are a Senior QA Engineer with 10+ years of experience.

Generate HIGH-QUALITY, NON-REPETITIVE, enterprise-level test cases for the following workflow.

WORKFLOW:
{workflow}


STRICT RULES:
- Do NOT create duplicate or repetitive scenario variations.
- Do NOT invent artificial cases (e.g. 'multiple clicks', 'same action different label').
- Focus on: business logic, real risk, edge cases, boundary values, negative testing, data validation.
- Group test cases logically by Module (e.g. Login, KPI Cards, Date Filter, Export, etc.).
- LIMIT: Generate 40–45 test cases MAXIMUM. Stop once all unique behaviors are covered. Do not pad.
- Each test case must add unique value that no other test case in the list covers.
- Include all scenario types: Positive (happy path), Negative (invalid input/errors), Boundary (min/max values), API/data validation.
- Write in a clean, human QA style — not like an AI checklist.

PRE-GENERATION CHECKLIST (do this mentally BEFORE writing):
1. List every unique feature in the workflow.
2. Map ONE test case per unique behavior.
3. If two test cases validate the same logic — discard one.
4. Do not rephrase the same scenario with different words.
5. Do not repeat the same Expected Result across different test cases.
6. Stop generating when all unique features are covered.

OUTPUT FORMAT RULES (STRICTLY FOLLOW):
- Each test case on its own line.
- Output EXACTLY 7 pipe "|" separated fields:
  TC_ID | Module | Test Scenario | Test Steps | Expected Result | Test Case Type | Priority
- TC_IDs must be unique and sequential: TC001, TC002, ...
- Module: Name of the feature/section being tested (e.g. Login, Revenue KPI Card, Date Filter, Export).
- Test Scenario: Concise, human-written title. Do NOT start with "Verify". Sound like a human QA wrote it.
- Test Steps:
    * Each step must start with an ACTION VERB: Enter, Click, Select, Verify, Observe, Open, Expand, Upload, Download, Clear, Set, Submit, Navigate, Enable, Disable, Scroll.
    * ONE action per step — never combine two actions into one step.
    * Separate ALL steps with semicolons (the system will automatically number them).
    * Do NOT write steps in paragraph or sentence form.
    * Do NOT start with 'Navigate to X' — navigation is implied from the scenario title.
    * FORBIDDEN: Combining actions e.g. 'Enter email and click Login' — these must be TWO separate steps.
    * FORBIDDEN: Listing multiple items in one step e.g. 'Verify Revenue Card, Orders Card, Trend Chart'.
    * Use actual element names from the workflow (exact button labels, field names, column headers).
- Expected Result:
    * Based on the FINAL actionable step only — not the setup or navigation.
    * ONE concise statement describing the system's response.
    * Do NOT repeat the same expected result across different test cases.
    * Rules by type:
        - Filter → data updates to show only matching records
        - Export → file downloads with correct name and format
        - Error → specific error message shown; action is blocked
        - Logout → session ends; user is redirected to login page
        - Form (valid) → record is saved; success confirmation shown
        - Form (invalid) → validation error displayed; form is not submitted
        - Delete → record removed from list; success notification shown
        - Search → only matching results displayed (or 'No results' if none)
- Test Case Type: EXACTLY one of — Functional, Negative, UI/UX, Validation
- Priority: EXACTLY one of — High, Medium, Low
- No header row. No markdown. No bullets. One test case per line only.

ONE COMPONENT = ONE TEST CASE:
- Each dashboard widget (Revenue Card, Orders Card, Line Chart, Bar Chart, Products Table) = its OWN test case.
- Never bundle multiple components into one test case.

EXAMPLE OUTPUT (notice: each step starts with an action verb, ONE action per step, steps separated by semicolons):
TC001 | Login | Successful login with valid credentials | Enter valid email in the Email field; Enter correct password in the Password field; Click the 'Login' button; Observe the dashboard page | Dashboard loads successfully and all modules are accessible from the sidebar | Functional | High
TC002 | Login | Login attempt with incorrect password | Enter a valid email in the Email field; Enter an incorrect password; Click the 'Login' button; Observe the error response | Error message 'Invalid credentials' appears; user remains on the login page | Negative | High
TC003 | Revenue KPI Card | Revenue KPI card displays correctly | Open the main dashboard; Observe the 'Total Revenue' KPI card; Verify the card shows a numeric currency value | Revenue KPI card is visible and displays a valid formatted monetary figure | Functional | High
TC004 | Date Filter | Apply a valid custom date range | Select the From Date as the first day of last month; Select the To Date as the last day of last month; Click the 'Apply' button; Observe the dashboard data | All charts and KPI cards refresh to show data only for the selected date range | Functional | High
TC005 | Date Filter | Validate From Date later than To Date | Set From Date to a future date; Set To Date to today's date; Click the 'Apply' button | Validation error 'From Date cannot be after To Date' is displayed; filter is not applied | Validation | High
TC006 | Export | Export currently filtered data to Excel | Select a region from the Region Filter dropdown; Click 'Apply'; Click the 'Export to Excel' button; Observe the file download | Excel file downloads with only the filtered records and correct column headers | Functional | Medium

Now generate {40}-{45} non-redundant enterprise-level test cases for the workflow above:"""


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

        # Split on at most 6 pipes so any '|' inside Expected Result doesn't
        # break the field count (the AI occasionally includes pipes in text).
        parts = [p.strip() for p in line.split("|", 6)]

        # Expect exactly 7 fields from the LLM:
        # TC_ID | Module | Scenario | Steps | Expected Result | Test Case Type | Priority
        if len(parts) != 7:
            print(
                f"[Groq] Warning: Skipping line {line_num} "
                f"(expected 7 fields, got {len(parts)}): {line!r}"
            )
            continue

        # Validate TC ID format (e.g. TC001)
        if not re.match(r"^TC\d+$", parts[0], re.IGNORECASE):
            print(
                f"[Groq] Warning: Skipping line {line_num} "
                f"(invalid TC ID '{parts[0]}'): {line!r}"
            )
            continue

        # Format steps (index 3) into a clean numbered list
        # Fields: parts[0]=TC_ID, parts[1]=Module, parts[2]=Scenario, parts[3]=Steps
        raw_steps = parts[3]
        step_items = [s.strip() for s in raw_steps.split(";") if s.strip()]

        # Auto-fix: if the AI collapsed many UI elements into one comma-list step
        # (e.g. "Check presence of X, Y, Z"), explode it into individual steps.
        expanded = []
        for step in step_items:
            # Detect a single step that is really a comma list of 3+ items
            # Pattern: starts with an action verb followed by a comma-separated list
            comma_parts = [c.strip() for c in step.split(",") if c.strip()]
            if len(comma_parts) >= 3 and any(
                step.lower().startswith(kw)
                for kw in ("check", "verify", "confirm", "ensure", "validate")
            ):
                # Keep the action verb from the first part and make each item its own step
                prefix_match = re.match(
                    r'^(check presence of|verify presence of|confirm|check|verify|ensure|validate)\s+',
                    step, re.IGNORECASE
                )
                prefix = prefix_match.group(0).rstrip() if prefix_match else "Confirm"
                for item in comma_parts:
                    # Remove the action prefix from the first item if it crept in
                    clean = re.sub(
                        r'^(check presence of|verify presence of|confirm|check|verify|ensure|validate)\s+',
                        '', item, flags=re.IGNORECASE
                    ).strip().rstrip(",")
                    if clean:
                        expanded.append(f"{prefix} '{clean}' is visible")
            else:
                expanded.append(step)

        if expanded:
            parts[3] = "\n".join(f"{i}. {step}" for i, step in enumerate(expanded, 1))

        # Sanitize Test Case Type (index 5) — pick the first recognised type word
        VALID_TC_TYPES = {"Functional", "Negative", "Positive", "UI/UX", "Validation", "Security"}
        tc_type_raw = parts[5]
        tc_type = next(
            (t for t in VALID_TC_TYPES if t.lower() in tc_type_raw.lower()),
            tc_type_raw.split()[0] if tc_type_raw.split() else "Functional",
        )
        parts[5] = tc_type

        # Sanitize Priority (index 6) — extract just High / Medium / Low
        VALID_PRIORITIES = {"High", "Medium", "Low"}
        priority_raw = parts[6]
        priority = next(
            (p for p in VALID_PRIORITIES if p.lower() in priority_raw.lower()),
            "Medium",
        )
        parts[6] = priority

        # Build the full 13-column row:
        # [TC_ID, Module, Scenario, Steps, Expected, <<Actual blank>>, TC_Type, Priority, <<Status blank>>, Date, <<Completed blank>>, <<Tested By blank>>, <<Comments blank>>]
        row = [
            parts[0],  # A: Test Case ID
            parts[1],  # B: Module
            parts[2],  # C: Test Scenario
            parts[3],  # D: Test Steps (numbered)
            parts[4],  # E: Expected Result
            "",        # F: Actual Result — blank for tester
            parts[5],  # G: Test Case Type — AI-generated
            parts[6],  # H: Priority — AI-generated
            "",        # I: Status — blank dropdown
            date.today().strftime("%d/%m/%Y"),  # J: Date
            "",        # K: Completed Date — blank for tester
            "",        # L: Tested By — blank for tester
            "",        # M: Comments — blank for tester
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
TC_TYPE_COL_INDEX = 6  # Column G (0-based), shifted right by 1 due to Module column B

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
