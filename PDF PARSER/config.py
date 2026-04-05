import os
from dotenv import load_dotenv

load_dotenv()


def _secret(key: str, default: str = "") -> str:
    """Читает секрет: сначала из st.secrets (Streamlit Cloud), потом из .env."""
    try:
        import streamlit as st
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return os.getenv(key, default)


OPENROUTER_API_KEY = _secret("OPENROUTER_API_KEY")
OPENROUTER_MODEL = _secret("OPENROUTER_MODEL") or "google/gemini-2.0-flash-001"
OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1"

GOOGLE_CREDENTIALS_FILE = _secret("GOOGLE_CREDENTIALS_FILE") or "credentials.json"
SPREADSHEET_ID = _secret("SPREADSHEET_ID")

PITCHES_FOLDER = _secret("PITCHES_FOLDER") or "./pitches"
STATE_FILE = _secret("STATE_FILE") or "./processed.json"

PDF_DPI = int(_secret("PDF_DPI") or "150")
