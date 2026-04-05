# Настройка PDF Pitch Parser

## 1. Установить зависимости системы

**macOS:**
```bash
brew install poppler
```

**Ubuntu/Debian:**
```bash
sudo apt-get install poppler-utils
```

---

## 2. Установить Python-пакеты

```bash
pip install -r requirements.txt
```

---

## 3. Настроить OpenRouter

1. Зарегистрироваться на [openrouter.ai](https://openrouter.ai)
2. Создать API-ключ в разделе Keys
3. Пополнить баланс (рекомендую $5–10 для старта)

Стоимость: Gemini 2.0 Flash ~$0.075 за 1M токенов ввода.
Один питч (15 слайдов) ≈ $0.01–0.03.

---

## 4. Настроить Google Sheets API

### Создать сервисный аккаунт:
1. Открыть [Google Cloud Console](https://console.cloud.google.com/)
2. Создать новый проект (или выбрать существующий)
3. Включить API: **Google Sheets API** и **Google Drive API**
4. Перейти в **IAM & Admin → Service Accounts → Create Service Account**
5. Скачать JSON-ключ → сохранить как `credentials.json` в папку проекта

### Создать таблицу:
1. Создать новую Google таблицу
2. Скопировать ID из URL: `https://docs.google.com/spreadsheets/d/`**`ВОТ_ЭТО_ID`**`/edit`
3. Поделиться таблицей с email сервисного аккаунта (из credentials.json поле `client_email`) — дать права **Редактора**

---

## 5. Создать .env файл

Скопировать `.env.example` в `.env` и заполнить:

```bash
cp .env.example .env
```

```
OPENROUTER_API_KEY=sk-or-v1-...
OPENROUTER_MODEL=google/gemini-2.0-flash-001
GOOGLE_CREDENTIALS_FILE=credentials.json
SPREADSHEET_ID=1ABC...XYZ
PITCHES_FOLDER=./pitches
STATE_FILE=./processed.json
PDF_DPI=150
```

---

## 6. Запуск

### Разовая обработка папки:
```bash
python main.py --meeting "Инвест встреча 05.04.2026"
```

### Указать другую папку:
```bash
python main.py --meeting "Инвест встреча 05.04.2026" --folder /path/to/pdfs
```

### Структура папок (рекомендуемая):
```
pitches/
├── pitch_startup1.pdf
├── pitch_startup2.pdf
└── pitch_startup3.pdf
```

---

## Как работает повторный запуск

Скрипт хранит список обработанных файлов в `processed.json`.
При повторном запуске — обрабатываются только **новые** PDF.
Чтобы перепарсить файл, удали его путь из `processed.json`.

---

## Структура таблицы

Каждая встреча → отдельный лист. Каждый питч → одна строка.

| Колонка | Описание |
|---|---|
| Название проекта | |
| Сфера / Индустрия | |
| Стадия | pre-seed / seed / раунд A |
| Проблема | |
| Решение | |
| Продукт | |
| Целевая аудитория | |
| Рынок (TAM/SAM/SOM) | |
| Бизнес-модель | |
| Трекшн / Метрики | |
| Команда | |
| Запрашиваемые инвестиции | |
| На что идут деньги | |
| Контакты | |
| Файл источника | имя PDF |
| Дата парсинга | |
