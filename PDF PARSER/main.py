#!/usr/bin/env python3
"""
PDF Pitch Parser
Парсит PDF-питчи с инвест-встреч и записывает данные в Google Sheets.

Использование:
  python main.py --meeting "Инвест встреча 05.04.2026"
  python main.py --meeting "Инвест встреча 05.04.2026" --folder ./pitches/april
"""
import argparse
import sys
from processor import process_folder


def main():
    parser = argparse.ArgumentParser(
        description="PDF Pitch Parser — извлекает данные из питчей в Google Sheets",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--meeting",
        required=True,
        help='Название встречи (станет именем листа в таблице). Пример: "Инвест встреча 05.04.2026"',
    )
    parser.add_argument(
        "--folder",
        default=None,
        help="Папка с PDF-файлами. Если не указана — берётся PITCHES_FOLDER из .env",
    )

    args = parser.parse_args()

    if not args.meeting.strip():
        print("Ошибка: название встречи не может быть пустым.")
        sys.exit(1)

    process_folder(meeting_name=args.meeting.strip(), folder=args.folder)


if __name__ == "__main__":
    main()
