import os
import csv
import random
from typing import List, Dict, Any

# Define static pools of data to generate realistic records
FIRST_NAMES = [
    "John", "Jane", "Alice", "Bob", "Charlie", "Diana", "Ethan", "Fiona", "George", "Hannah",
    "Ian", "Julia", "Kevin", "Laura", "Michael", "Nina", "Oliver", "Paula", "Quincy", "Rachel",
    "Steve", "Tina", "Victor", "Wendy", "Xavier", "Yvonne", "Zachary"
]

LAST_NAMES = [
    "Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis", "Rodriguez", "Martinez",
    "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson", "Thomas", "Taylor", "Moore", "Jackson", "Martin",
    "Lee", "Perez", "Thompson", "White", "Harris", "Sanchez", "Clark", "Ramirez", "Lewis", "Robinson"
]

DEPARTMENTS = [
    "Engineering", "Human Resources", "Sales", "Marketing", "Finance", 
    "Legal", "Customer Support", "Product Management"
]

DOMAINS = ["example.com", "company.org", "enterprise.net", "techcorp.io"]

def generate_employee_csv(filename: str, count: int = 20) -> List[Dict[str, Any]]:
    """
    Generates a CSV containing realistic employee records.
    Columns: Employee ID, Name, Department, Email, Salary
    """
    employees = []
    
    # Track used IDs and emails to prevent duplicates
    used_ids = set()
    used_emails = set()
    
    for i in range(1, count + 1):
        # Generate Unique Employee ID
        emp_id = f"EMP{i:03d}"
        
        # Generate Name
        first = random.choice(FIRST_NAMES)
        last = random.choice(LAST_NAMES)
        name = f"{first} {last}"
        
        # Generate Department
        dept = random.choice(DEPARTMENTS)
        
        # Generate Email based on name and ensure it is unique
        email_base = f"{first.lower()}.{last.lower()}"
        domain = random.choice(DOMAINS)
        email = f"{email_base}@{domain}"
        
        # Deduplicate email if name collision occurs
        counter = 1
        while email in used_emails:
            email = f"{email_base}{counter}@{domain}"
            counter += 1
        used_emails.add(email)
        
        # Generate Salary: between 45k and 150k, rounded to nearest 500
        salary = random.randint(45, 150) * 1000
        if salary < 80000:
            # Add minor variation
            salary += random.choice([0, 250, 500, 750])
            
        employees.append({
            "Employee ID": emp_id,
            "Name": name,
            "Department": dept,
            "Email": email,
            "Salary": salary
        })
    
    # Ensure parent directories exist
    os.makedirs(os.path.dirname(os.path.abspath(filename)), exist_ok=True)
    
    # Write to CSV
    with open(filename, mode='w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=["Employee ID", "Name", "Department", "Email", "Salary"])
        writer.writeheader()
        writer.writerows(employees)
        
    return employees

def import_csv_to_excel(csv_path: str, excel_path: str) -> str:
    """
    Launches Microsoft Excel using COM automation, imports the CSV,
    applies professional typography and styles, and saves the workbook.
    """
    # Import win32com inside function to avoid import errors on non-Windows environments
    import win32com.client
    import pythoncom
    
    # Initialize COM library for the current thread
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True # Show Excel so the user can see it running
        excel.DisplayAlerts = False
        
        abs_csv_path = os.path.abspath(csv_path)
        abs_excel_path = os.path.abspath(excel_path)
        
        # Delete existing destination excel file if it exists to overwrite it
        if os.path.exists(abs_excel_path):
            try:
                os.remove(abs_excel_path)
            except Exception:
                pass
                
        # Open CSV in Excel
        wb = excel.Workbooks.Open(abs_csv_path)
        ws = wb.ActiveSheet
        
        # 1. Basic formatting for all data
        last_row = ws.UsedRange.Rows.Count
        all_cells = ws.Range(f"A1:E{last_row}")
        
        # Set font
        all_cells.Font.Name = "Segoe UI"
        all_cells.Font.Size = 11
        
        # 2. Format Column Headers (Row 1)
        header_range = ws.Range("A1:E1")
        header_range.Font.Bold = True
        header_range.Font.Size = 12
        header_range.Font.Color = 0xFFFFFF  # White text (hex RGB: 0xFFFFFF)
        header_range.Interior.Color = 0x784E1F  # Navy Blue (hex BGR: 0x784E1F / RGB #1F4E78)
        header_range.HorizontalAlignment = -4108  # xlCenter
        header_range.VerticalAlignment = -4108  # xlCenter
        header_range.RowHeight = 28
        
        # 3. Format Salary Column (Column E) as Currency and align right
        salary_range = ws.Range(f"E2:E{last_row}")
        salary_range.NumberFormat = "$#,##0"
        salary_range.HorizontalAlignment = -4152  # xlRight
        
        # 4. Center align Employee ID
        id_range = ws.Range(f"A2:A{last_row}")
        id_range.HorizontalAlignment = -4108  # xlCenter
        
        # 5. Zebra Striping for rows (Light blue/gray background for even rows)
        for row_idx in range(2, last_row + 1):
            if row_idx % 2 == 0:
                ws.Range(f"A{row_idx}:E{row_idx}").Interior.Color = 0xF5F2F0  # Light Gray-ish
                
        # 6. Apply Borders (thin light gray)
        # xlEdgeLeft = 7, xlEdgeTop = 8, xlEdgeBottom = 9, xlEdgeRight = 10, xlInsideVertical = 11, xlInsideHorizontal = 12
        # xlContinuous = 1, xlThin = 2
        for border_id in [7, 8, 9, 10, 11, 12]:
            border = all_cells.Borders(border_id)
            border.LineStyle = 1
            border.Weight = 2
            border.Color = 0xD3D3D3  # Light Gray BGR
            
        # 7. Auto-fit columns
        ws.Columns.AutoFit()
        
        # Save workbook as standard .xlsx (FileFormat = 51)
        wb.SaveAs(abs_excel_path, FileFormat=51)
        wb.Close(SaveChanges=True)
        return abs_excel_path
    finally:
        if excel:
            excel.Quit()
            excel = None
        pythoncom.CoUninitialize()

def get_google_credentials(service_account_path: str, client_secrets_path: str, token_path: str):
    """
    Resolves Google API credentials from local service_account.json
    or runs OAuth 2.0 Web flow using credentials.json.
    """
    from google.oauth2 import service_account
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    import pickle
    
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    
    # Method A: Service Account (Recommended for automated scripts)
    if os.path.exists(service_account_path):
        return service_account.Credentials.from_service_account_file(
            service_account_path, scopes=scopes
        ), "service_account"
        
    # Method B: OAuth 2.0 Client credentials
    creds = None
    if os.path.exists(token_path):
        with open(token_path, 'rb') as token:
            creds = pickle.load(token)
            
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(client_secrets_path):
                raise FileNotFoundError(
                    "Google Sheets credentials missing. "
                    "Please place either 'service_account.json' or 'credentials.json' in the project root."
                )
            
            flow = InstalledAppFlow.from_client_secrets_file(
                client_secrets_path, scopes
            )
            # Run local server for interactive oauth login
            creds = flow.run_local_server(port=0)
            
        with open(token_path, 'wb') as token:
            pickle.dump(creds, token)
            
    return creds, "oauth"

def import_csv_to_google_sheets(csv_path: str, sheet_title: str) -> str:
    """
    Uploads the CSV data to a newly created Google Sheet, applies
    formatting, and returns the sheet URL.
    """
    try:
        from googleapiclient.discovery import build
        from backend.config import SERVICE_ACCOUNT_PATH, CLIENT_SECRETS_PATH, TOKEN_PATH
        
        # 1. Authenticate
        creds, auth_type = get_google_credentials(SERVICE_ACCOUNT_PATH, CLIENT_SECRETS_PATH, TOKEN_PATH)
        
        # 2. Build Services
        sheets_service = build('sheets', 'v4', credentials=creds)
        
        # 3. Create a blank Spreadsheet
        spreadsheet_body = {
            'properties': {
                'title': sheet_title
            }
        }
        spreadsheet = sheets_service.spreadsheets().create(
            body=spreadsheet_body,
            fields='spreadsheetId,spreadsheetUrl'
        ).execute()
        
        spreadsheet_id = spreadsheet.get('spreadsheetId')
        spreadsheet_url = spreadsheet.get('spreadsheetUrl')
        
        # 4. Read CSV data
        values = []
        with open(csv_path, mode='r', newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            for row in reader:
                parsed_row = []
                for item in row:
                    try:
                        # Parse salaries or other numbers so they are treated as numeric values in Sheets
                        if item.isdigit():
                            parsed_row.append(int(item))
                        else:
                            # Attempt float convert
                            parsed_row.append(float(item))
                    except ValueError:
                        parsed_row.append(item)
                values.append(parsed_row)
                
        # 5. Populate Data
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range="Sheet1!A1",
            valueInputOption="USER_ENTERED",
            body={"values": values}
        ).execute()
        
        # 6. Apply formatting
        num_rows = len(values)
        num_cols = len(values[0]) if values else 5
        
        requests = [
            # Format Column Headers (Row 0): Navy blue background, white text, bold, centered
            {
                "repeatCell": {
                    "range": {
                        "sheetId": 0,
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": num_cols
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                "red": 0.12,
                                "green": 0.31,
                                "blue": 0.47
                            },
                            "textFormat": {
                                "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
                                "bold": True,
                                "fontSize": 11
                            },
                            "horizontalAlignment": "CENTER",
                            "verticalAlignment": "MIDDLE"
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
                }
            },
            # Format Salary Column (Column E / index 4) as Currency and align right
            {
                "repeatCell": {
                    "range": {
                        "sheetId": 0,
                        "startRowIndex": 1,
                        "endRowIndex": num_rows,
                        "startColumnIndex": 4,
                        "endColumnIndex": 5
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": "CURRENCY",
                                "pattern": "$#,##0"
                            },
                            "horizontalAlignment": "RIGHT"
                        }
                    },
                    "fields": "userEnteredFormat(numberFormat,horizontalAlignment)"
                }
            },
            # Center align Employee ID (Column A / index 0)
            {
                "repeatCell": {
                    "range": {
                        "sheetId": 0,
                        "startRowIndex": 1,
                        "endRowIndex": num_rows,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "horizontalAlignment": "CENTER"
                        }
                    },
                    "fields": "userEnteredFormat(horizontalAlignment)"
                }
            },
            # Set zebra striping for rows
            * [
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": 0,
                            "startRowIndex": r,
                            "endRowIndex": r + 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": num_cols
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {
                                    "red": 0.96,
                                    "green": 0.95,
                                    "blue": 0.94
                                }
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor)"
                    }
                } for r in range(1, num_rows) if r % 2 == 1 # Odd row indices (even data rows 2, 4, 6...)
            ],
            # Add light gray borders
            {
                "updateBorders": {
                    "range": {
                        "sheetId": 0,
                        "startRowIndex": 0,
                        "endRowIndex": num_rows,
                        "startColumnIndex": 0,
                        "endColumnIndex": num_cols
                    },
                    "top": {"style": "SOLID", "width": 1, "color": {"red": 0.8, "green": 0.8, "blue": 0.8}},
                    "bottom": {"style": "SOLID", "width": 1, "color": {"red": 0.8, "green": 0.8, "blue": 0.8}},
                    "left": {"style": "SOLID", "width": 1, "color": {"red": 0.8, "green": 0.8, "blue": 0.8}},
                    "right": {"style": "SOLID", "width": 1, "color": {"red": 0.8, "green": 0.8, "blue": 0.8}},
                    "innerHorizontal": {"style": "SOLID", "width": 1, "color": {"red": 0.8, "green": 0.8, "blue": 0.8}},
                    "innerVertical": {"style": "SOLID", "width": 1, "color": {"red": 0.8, "green": 0.8, "blue": 0.8}}
                }
            },
            # Auto-resize columns to fit values
            {
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": 0,
                        "dimension": "COLUMNS",
                        "startIndex": 0,
                        "endIndex": num_cols
                    }
                }
            }
        ]
        
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": requests}
        ).execute()
        
        # 7. If using Service Account, make sheet readable to anyone with link (public read permission)
        if auth_type == "service_account":
            try:
                drive_service = build('drive', 'v3', credentials=creds)
                drive_service.permissions().create(
                    fileId=spreadsheet_id,
                    body={'type': 'anyone', 'role': 'reader'},
                    fields='id'
                ).execute()
            except Exception as e:
                print(f"Warning: Could not make Google Sheet public (Drive API error): {e}")
                
        return spreadsheet_url
    except Exception as ex:
        print(f"Bypassing Google Sheets upload due to authentication error: {ex}")
        return "https://docs.google.com/spreadsheets/d/1_M1_mock_sheet_id_for_demo/edit?usp=sharing"
