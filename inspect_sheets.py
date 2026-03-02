import os
import gspread
from dotenv import load_dotenv

load_dotenv()

def main():
    gc = gspread.service_account(filename="service_account.json")
    sheet_id = os.environ.get("GOOGLE_SHEET_ID")
    print(f"Sheet ID: {sheet_id}")
    spreadsheet = gc.open_by_key(sheet_id)
    
    # Get the spreadsheet using the raw API to see all properties
    res = spreadsheet.client.request(
        'get',
        f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}?includeGridData=true"
    )
    
    data = res.json()
    sheet = data['sheets'][0]
    grid_data = sheet.get('data', [])[0]
    
    # Let's inspect row 2, column F (index 5) depending on where the user's dropdowns are
    # In their screenshot, Test Case Type is column C or E? In the latest it's F.
    # We will just print the dataValidation for the first 10 columns of row 2
    row_data = grid_data.get('rowData', [])
    if len(row_data) > 1:
        cells = row_data[1].get('values', [])
        for i, cell in enumerate(cells):
            if 'dataValidation' in cell:
                print(f"Column {i} Data Validation:")
                print(cell['dataValidation'])

if __name__ == "__main__":
    main()
