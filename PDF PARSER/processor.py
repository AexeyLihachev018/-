import json
import os
from pathlib import Path

from config import PITCHES_FOLDER, STATE_FILE, PDF_DPI
from pdf_to_images import pdf_to_base64_images
from extractor import extract_pitch_data
from sheets import get_or_create_sheet, write_pitch


def load_state() -> set:
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, encoding="utf-8") as f:
            return set(json.load(f))
    return set()


def save_state(processed: set):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(sorted(processed), f, ensure_ascii=False, indent=2)


def process_folder(meeting_name: str, folder: str = None):
    """
    Обрабатывает все новые PDF в папке и записывает данные в Google Sheets.

    Args:
        meeting_name: Название встречи — станет именем листа в таблице.
        folder: Путь к папке с PDF. Если не указан — берётся из .env (PITCHES_FOLDER).
    """
    folder_path = Path(folder) if folder else Path(PITCHES_FOLDER)

    if not folder_path.exists():
        print(f"Папка не найдена: {folder_path}")
        return

    processed = load_state()
    pdf_files = sorted(folder_path.glob("*.pdf"))
    new_files = [f for f in pdf_files if str(f.resolve()) not in processed]

    if not new_files:
        print(f"Нет новых PDF в папке: {folder_path}")
        print(f"Всего файлов в папке: {len(pdf_files)}, уже обработано: {len(processed)}")
        return

    print(f"Найдено новых PDF: {len(new_files)}")
    sheet = get_or_create_sheet(meeting_name)

    success_count = 0
    error_count = 0

    for pdf_path in new_files:
        print(f"\n[{new_files.index(pdf_path) + 1}/{len(new_files)}] {pdf_path.name}")
        try:
            print("  → Конвертирую страницы в изображения...")
            images = pdf_to_base64_images(str(pdf_path), dpi=PDF_DPI)
            print(f"  → {len(images)} стр., отправляю в модель...")

            data = extract_pitch_data(images)
            project_name = data.get("название_проекта", "?")

            write_pitch(sheet, data, pdf_path.name)
            processed.add(str(pdf_path.resolve()))
            save_state(processed)

            print(f"  ✓ Записано: «{project_name}»")
            success_count += 1

        except Exception as e:
            print(f"  ✗ Ошибка: {e}")
            error_count += 1

    print(f"\n{'='*40}")
    print(f"Готово. Успешно: {success_count}, ошибок: {error_count}")
    if success_count > 0:
        print(f"Данные записаны на лист «{meeting_name}»")
