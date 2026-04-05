# PDF Pitch Parser — Руководство по запуску

## Что нужно (одноразовая настройка)

### 1. Установить Python-пакеты
```bash
pip install -r requirements.txt
```

### 2. Файл `.env` — должен лежать в папке проекта
```
OPENROUTER_API_KEY=sk-or-v1-...
OPENROUTER_MODEL=google/gemini-2.0-flash-001
GOOGLE_CREDENTIALS_FILE=credentials.json
SPREADSHEET_ID=1ABC...XYZ
PITCHES_FOLDER=./pitches
STATE_FILE=./processed.json
PDF_DPI=150
```

### 3. Файл `credentials.json` — должен лежать в папке проекта
Скачивается из Google Cloud Console → IAM → Service Accounts → Keys.

### 4. Google Таблица — дать доступ сервисному аккаунту
Открыть таблицу → Поделиться → добавить email из поля `client_email` в `credentials.json` → права Редактора.

---

## Как запускать парсинг (каждый раз)

### Через веб-интерфейс (рекомендуется)
Дважды кликнуть на файл **`Запустить приложение.command`** → откроется браузер на localhost:8501.

Или из терминала:
```bash
python3 -m streamlit run app.py
```

### Через терминал (альтернатива)
```bash
python main.py --meeting "Инвест встреча 05.04.2026"
```

---

## Стоимость
Gemini 2.0 Flash: ~$0.01–0.03 за один питч (15 слайдов).
