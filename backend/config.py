import os
from pathlib import Path
from dotenv import load_dotenv

# Find project root
BASE_DIR = Path(__file__).resolve().parent.parent

# Load environment variables
load_dotenv(BASE_DIR / ".env")

# API Keys & Configurations
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
PORT = int(os.getenv("PORT", "8000"))
HOST = os.getenv("HOST", "127.0.0.1")

# File Outputs (absolute paths to prevent working directory issues)
CSV_OUTPUT_PATH = str(BASE_DIR / "employees.csv")
EXCEL_OUTPUT_PATH = str(BASE_DIR / os.getenv("EXCEL_OUTPUT_PATH", "employees.xlsx"))
GOOGLE_SHEET_TITLE = os.getenv("GOOGLE_SHEET_TITLE", "Employee CSV Import")

# Google API Paths
SERVICE_ACCOUNT_PATH = str(BASE_DIR / "service_account.json")
CLIENT_SECRETS_PATH = str(BASE_DIR / "credentials.json")
TOKEN_PATH = str(BASE_DIR / "token.json")
