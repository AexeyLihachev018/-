import os
from dotenv import load_dotenv

load_dotenv()

OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")
OPENROUTER_MODEL = os.getenv("OPENROUTER_MODEL", "google/gemini-2.0-flash-001")
OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1"

GOOGLE_CREDENTIALS_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE", "credentials.json")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")

PITCHES_FOLDER = os.getenv("PITCHES_FOLDER", "./pitches")
STATE_FILE = os.getenv("STATE_FILE", "./processed.json")

PDF_DPI = int(os.getenv("PDF_DPI", "150"))
