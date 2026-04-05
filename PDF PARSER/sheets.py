from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
from config import GOOGLE_CREDENTIALS_FILE, SPREADSHEET_ID

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

COLUMNS = [
    "Название проекта",
    "Сфера / Индустрия",
    "Стадия",
    "Проблема",
    "Решение",
    "Продукт",
    "Целевая аудитория",
    "Рынок (TAM/SAM/SOM)",
    "Бизнес-модель",
    "Трекшн / Метрики",
    "Команда",
    "Запрашиваемые инвестиции",
    "На что идут деньги",
    "Контакты",
    "Оценка (1-10)",
    "Комментарий инвестора",
    "Файл источника",
    "Дата парсинга",
]

# Маппинг колонок на ключи JSON от модели
FIELD_MAP = {
    "Название проекта": "название_проекта",
    "Сфера / Индустрия": "сфера_индустрия",
    "Стадия": "стадия",
    "Проблема": "проблема",
    "Решение": "решение",
    "Продукт": "продукт",
    "Целевая аудитория": "целевая_аудитория",
    "Рынок (TAM/SAM/SOM)": "рынок",
    "Бизнес-модель": "бизнес_модель",
    "Трекшн / Метрики": "тракшн_метрики",
    "Команда": "команда",
    "Запрашиваемые инвестиции": "инвестиции_запрос",
    "На что идут деньги": "бюджет_на_что",
    "Контакты": "контакты",
    "Оценка (1-10)": "оценка_привлекательности",
    "Комментарий инвестора": "комментарий_инвестора",
}


def _get_client() -> gspread.Client:
    # На Streamlit Cloud — читаем из st.secrets
    try:
        import streamlit as st
        if "gcp_service_account" in st.secrets:
            creds = Credentials.from_service_account_info(
                dict(st.secrets["gcp_service_account"]),
                scopes=SCOPES,
            )
            return gspread.authorize(creds)
    except Exception:
        pass
    # Локально — читаем из файла credentials.json
    creds = Credentials.from_service_account_file(GOOGLE_CREDENTIALS_FILE, scopes=SCOPES)
    return gspread.authorize(creds)


def get_or_create_sheet(meeting_name: str) -> gspread.Worksheet:
    """Возвращает лист встречи, создаёт если не существует."""
    client = _get_client()
    spreadsheet = client.open_by_key(SPREADSHEET_ID)

    try:
        sheet = spreadsheet.worksheet(meeting_name)
        print(f"Лист '{meeting_name}' найден.")
    except gspread.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(
            title=meeting_name,
            rows=1000,
            cols=len(COLUMNS),
        )
        sheet.append_row(COLUMNS)
        # Жирный заголовок с синим фоном
        last_col_letter = chr(ord("A") + len(COLUMNS) - 1)
        sheet.format(f"A1:{last_col_letter}1", {
            "textFormat": {"bold": True},
            "backgroundColor": {"red": 0.85, "green": 0.92, "blue": 0.98},
        })
        print(f"Лист '{meeting_name}' создан.")

    return sheet


def is_duplicate(sheet: gspread.Worksheet, filename: str) -> bool:
    """Проверяет, был ли файл уже записан в этот лист."""
    col_index = COLUMNS.index("Файл источника") + 1  # 1-based
    existing_files = sheet.col_values(col_index)
    return filename in existing_files


def write_pitch(sheet: gspread.Worksheet, data: dict, filename: str) -> bool:
    """
    Добавляет строку с данными питча в лист.
    Возвращает True если записано, False если дубль (файл уже есть в листе).
    """
    if is_duplicate(sheet, filename):
        return False

    row = []
    for col in COLUMNS:
        if col == "Файл источника":
            row.append(filename)
        elif col == "Дата парсинга":
            row.append(datetime.now().strftime("%d.%m.%Y %H:%M"))
        else:
            field_key = FIELD_MAP.get(col, "")
            row.append(data.get(field_key, "не указано"))

    sheet.append_row(row, value_input_option="USER_ENTERED")
    return True
