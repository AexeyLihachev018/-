import os
import tempfile
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv

load_dotenv()

from pdf_to_images import pdf_to_base64_images
from extractor import extract_pitch_data
from sheets import get_or_create_sheet, write_pitch

# ── Настройки страницы ──────────────────────────────────────────────────────
st.set_page_config(
    page_title="Pitch Deck Parser",
    page_icon="📊",
    layout="wide",
)

# ── Стили ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
.metric-card {
    background: #1e2130;
    border-radius: 10px;
    padding: 16px 20px;
    margin: 4px 0;
}
.success-badge {
    background: #1a4731;
    color: #4ade80;
    padding: 3px 10px;
    border-radius: 20px;
    font-size: 13px;
}
.error-badge {
    background: #4c1d1d;
    color: #f87171;
    padding: 3px 10px;
    border-radius: 20px;
    font-size: 13px;
}
</style>
""", unsafe_allow_html=True)

# ── Sidebar ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Настройки")
    st.divider()

    api_key = os.getenv("OPENROUTER_API_KEY", "")
    model = os.getenv("OPENROUTER_MODEL", "google/gemini-2.0-flash-001")
    spreadsheet_id = os.getenv("SPREADSHEET_ID", "")

    st.text_input("OpenRouter API Key",
                  value=api_key[:20] + "..." if len(api_key) > 20 else api_key,
                  disabled=True)
    st.text_input("Модель", value=model, disabled=True)
    st.text_input("Google Sheets ID",
                  value=spreadsheet_id[:20] + "..." if len(spreadsheet_id) > 20 else spreadsheet_id,
                  disabled=True)

    st.divider()
    st.markdown("**Извлекаемые блоки:**")
    st.markdown("""
- Название проекта
- Сфера / Индустрия
- Стадия
- Проблема
- Решение
- Продукт
- Целевая аудитория
- Рынок (TAM/SAM/SOM)
- Бизнес-модель
- Трекшн / Метрики
- Команда
- Инвестиции
- Бюджет
- Контакты
""")

# ── Главная область ──────────────────────────────────────────────────────────
st.markdown("# 📊 Pitch Deck Parser")
st.markdown("Загрузите PDF питч-деки — приложение извлечёт ключевую информацию и отправит в Google Sheets.")
st.divider()

col1, col2 = st.columns([2, 1])

with col1:
    meeting_name = st.text_input(
        "🗓️ Название встречи",
        placeholder="Инвест встреча 05.04.2026",
        help="Станет именем листа в Google Sheets"
    )

with col2:
    dpi = st.select_slider(
        "Качество сканирования",
        options=[100, 150, 200],
        value=150,
        help="Выше = точнее, но дольше и дороже"
    )

uploaded_files = st.file_uploader(
    "Загрузите PDF файлы",
    type=["pdf"],
    accept_multiple_files=True,
    help="Можно загружать несколько файлов сразу"
)

if uploaded_files:
    est_sec = len(uploaded_files) * 20
    st.info(f"Загружено файлов: **{len(uploaded_files)}**  •  Примерное время: **~{est_sec} сек.**")

st.divider()

# ── Кнопка запуска ───────────────────────────────────────────────────────────
run = st.button(
    "🚀 Запустить парсинг",
    type="primary",
    disabled=not (uploaded_files and meeting_name.strip()),
    use_container_width=True,
)

if not meeting_name.strip() and uploaded_files:
    st.warning("Введите название встречи чтобы запустить парсинг")

# ── Обработка ────────────────────────────────────────────────────────────────
if run and uploaded_files and meeting_name.strip():

    sheet = None
    results = []

    progress_bar = st.progress(0, text="Подготовка...")
    status_container = st.container()

    with status_container:
        for i, uploaded_file in enumerate(uploaded_files):
            pct = int(i / len(uploaded_files) * 100)
            progress_bar.progress(pct, text=f"Обрабатываю {uploaded_file.name}...")

            with st.status(f"📄 {uploaded_file.name}", expanded=True) as file_status:
                try:
                    # Сохраняем временный файл
                    st.write("Конвертирую страницы в изображения...")
                    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                        tmp.write(uploaded_file.read())
                        tmp_path = tmp.name

                    images = pdf_to_base64_images(tmp_path, dpi=dpi)
                    os.unlink(tmp_path)

                    st.write(f"Отправляю {len(images)} стр. в модель...")
                    data = extract_pitch_data(images)

                    # Записываем в Sheets
                    if sheet is None:
                        st.write("Подключаюсь к Google Sheets...")
                        sheet = get_or_create_sheet(meeting_name.strip())

                    written = write_pitch(sheet, data, uploaded_file.name)

                    project_name = data.get("название_проекта", "?")
                    if written:
                        file_status.update(label=f"✅ {uploaded_file.name} → «{project_name}»", state="complete")
                        results.append({"file": uploaded_file.name, "project": project_name, "ok": True, "duplicate": False, "data": data})
                    else:
                        file_status.update(label=f"⚠️ {uploaded_file.name} — уже в таблице (пропущен)", state="complete")
                        st.warning("Файл уже был записан в этот лист ранее. Дубль не создан.")
                        results.append({"file": uploaded_file.name, "project": project_name, "ok": True, "duplicate": True, "data": data})

                except Exception as e:
                    file_status.update(label=f"❌ {uploaded_file.name} — ошибка", state="error")
                    st.error(str(e))
                    results.append({"file": uploaded_file.name, "project": "—", "ok": False, "data": {}})

    progress_bar.progress(100, text="Готово!")

    # ── Итоги ────────────────────────────────────────────────────────────────
    st.divider()
    ok_count = sum(1 for r in results if r["ok"] and not r.get("duplicate"))
    dup_count = sum(1 for r in results if r.get("duplicate"))
    err_count = len(results) - ok_count - dup_count

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Обработано файлов", len(results))
    col2.metric("Записано", ok_count)
    col3.metric("Дублей пропущено", dup_count)
    col4.metric("Ошибок", err_count)

    if ok_count > 0 or dup_count > 0:
        sheets_url = f"https://docs.google.com/spreadsheets/d/{os.getenv('SPREADSHEET_ID')}/edit"
        if ok_count > 0:
            st.success(f"Данные записаны на лист **«{meeting_name}»**")
        if dup_count > 0:
            st.info(f"Пропущено дублей: **{dup_count}** — эти файлы уже были в таблице")
        st.link_button("📊 Открыть Google Sheets", sheets_url, use_container_width=True)

    # ── Таблица результатов ───────────────────────────────────────────────────
    if results:
        st.divider()
        st.markdown("### Результаты парсинга")

        for r in results:
            if not r["ok"]:
                continue
            d = r["data"]
            score = d.get("оценка_привлекательности", "—")
            dup_label = " *(дубль, не записан)*" if r.get("duplicate") else ""
            expander_label = f"📄 {r['file']}  →  **{r['project']}**  |  Оценка: **{score}/10**{dup_label}"
            with st.expander(expander_label):
                # Блок оценки
                comment = d.get("комментарий_инвестора", "—")
                st.info(f"**Комментарий инвестора:** {comment}")

                col1, col2 = st.columns(2)
                with col1:
                    st.markdown(f"**Сфера:** {d.get('сфера_индустрия', '—')}")
                    st.markdown(f"**Стадия:** {d.get('стадия', '—')}")
                    st.markdown(f"**Инвестиции:** {d.get('инвестиции_запрос', '—')}")
                    st.markdown(f"**Команда:** {d.get('команда', '—')}")
                with col2:
                    st.markdown(f"**Проблема:** {d.get('проблема', '—')}")
                    st.markdown(f"**Решение:** {d.get('решение', '—')}")
                    st.markdown(f"**Рынок:** {d.get('рынок', '—')}")
                    st.markdown(f"**Контакты:** {d.get('контакты', '—')}")
