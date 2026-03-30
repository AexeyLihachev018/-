// ============================================================
//  TELEGRAM BOT — Отчёт РК Маркетинг
//  Ежедневный сбор данных по источникам трафика → Google Sheets
// ============================================================
//
//  НАСТРОЙКА (заполни перед запуском):
//  1. BOT_TOKEN  — получи у @BotFather в Telegram
//  2. CHAT_ID    — твой Telegram chat_id (узнай через @userinfobot)
//  3. SHEET_ID   — ID твоей Google таблицы (из URL)
//
//  ПЕРВЫЙ ЗАПУСК:
//  1. Вставь код в Apps Script
//  2. Заполни константы ниже
//  3. Запусти setupBot() — один раз
//  4. Запусти setupDailyTrigger() — один раз
//
// ============================================================

const BOT_TOKEN   = 'ВСТАВЬ_ТОКЕН_БОТА';       // от @BotFather
const CHAT_ID     = 'ВСТАВЬ_СВОЙ_CHAT_ID';     // твой Telegram ID
const SHEET_ID    = 'ВСТАВЬ_ID_ТАБЛИЦЫ';       // из URL Google Sheets
const DATA_SHEET  = 'Данные бота';              // название листа для хранения

// Источники трафика (порядок = порядок вопросов)
const SOURCES = [
  'Яндекс Директ',
  'SEO',
  'Почта',
  'Звонки',
  'Старые клиенты',
  'ФОС / Carrotquest',
  'Прочий трафик'
];

// ============================================================
//  WEBHOOK — обработчик входящих сообщений от Telegram
// ============================================================

function doPost(e) {
  try {
    const update = JSON.parse(e.postData.contents);
    if (!update.message) return;

    const chatId = String(update.message.chat.id);
    const text   = (update.message.text || '').trim();

    // Принимаем сообщения только от разрешённого пользователя
    if (chatId !== String(CHAT_ID)) {
      sendMessage(chatId, '⛔ У вас нет доступа к этому боту.');
      return;
    }

    handleMessage(chatId, text);

  } catch (err) {
    Logger.log('doPost error: ' + err.message);
  }
}

// ============================================================
//  ЛОГИКА ДИАЛОГА
// ============================================================

function handleMessage(chatId, text) {
  const props = PropertiesService.getUserProperties();
  const state = parseInt(props.getProperty('state') || '0');

  // Команды управления
  if (text === '/start' || text === '/начать') {
    startDialog(chatId, props);
    return;
  }

  if (text === '/отмена' || text === '/cancel') {
    props.deleteAllProperties();
    sendMessage(chatId, '❌ Ввод данных отменён. Напиши /начать чтобы начать заново.');
    return;
  }

  if (text === '/статус') {
    showStatus(chatId);
    return;
  }

  if (text === '/помощь' || text === '/help') {
    sendHelp(chatId);
    return;
  }

  // Диалог сбора данных
  if (state >= 1 && state <= SOURCES.length) {
    processDataInput(chatId, text, state, props);
    return;
  }

  // Если нет активного диалога
  sendMessage(chatId,
    '👋 Привет!\n\n' +
    'Напиши /начать чтобы внести данные за сегодня.\n' +
    'Или /помощь для списка команд.'
  );
}

function startDialog(chatId, props) {
  props.setProperty('state', '1');
  props.setProperty('data', JSON.stringify({}));

  const today = formatDate(new Date());
  const weekNum = getWeekOfMonth(new Date());

  sendMessage(chatId,
    '📊 *Отчёт РК Маркетинг*\n' +
    '📅 Дата: ' + today + ' (Неделя ' + weekNum + ')\n\n' +
    'Буду задавать вопросы по каждому источнику трафика.\n' +
    'Вводи только цифры (количество обращений).\n' +
    'Для отмены: /отмена\n\n' +
    '━━━━━━━━━━━━━━━━━━\n' +
    '1️⃣ *' + SOURCES[0] + '*\n' +
    'Сколько обращений пришло сегодня?'
  );
}

function processDataInput(chatId, text, state, props) {
  // Проверяем что введено число
  const value = parseInt(text);
  if (isNaN(value) || value < 0) {
    sendMessage(chatId,
      '⚠️ Введи только число (например: 5 или 0)\n\n' +
      'Вопрос повторяю:\n' +
      getEmojiNum(state) + ' *' + SOURCES[state - 1] + '*\n' +
      'Сколько обращений?'
    );
    return;
  }

  // Сохраняем ответ
  const data = JSON.parse(props.getProperty('data') || '{}');
  data[SOURCES[state - 1]] = value;
  props.setProperty('data', JSON.stringify(data));

  // Следующий вопрос или финал
  if (state < SOURCES.length) {
    const nextState = state + 1;
    props.setProperty('state', String(nextState));

    sendMessage(chatId,
      '✅ Записал: ' + SOURCES[state - 1] + ' = *' + value + '*\n\n' +
      '━━━━━━━━━━━━━━━━━━\n' +
      getEmojiNum(nextState) + ' *' + SOURCES[nextState - 1] + '*\n' +
      'Сколько обращений?'
    );

  } else {
    // Все данные собраны — записываем в таблицу
    props.setProperty('state', '0');
    data[SOURCES[state - 1]] = value;
    saveToSheet(chatId, data, props);
  }
}

function saveToSheet(chatId, data, props) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    let sheet   = ss.getSheetByName(DATA_SHEET);

    // Создаём лист если нет
    if (!sheet) {
      sheet = ss.insertSheet(DATA_SHEET);
      // Заголовки
      const headers = ['Дата', 'Неделя', 'Месяц', 'Год'].concat(SOURCES).concat(['ИТОГО']);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#4a4a4a')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    const now     = new Date();
    const total   = SOURCES.reduce((sum, s) => sum + (data[s] || 0), 0);
    const weekNum = getWeekOfMonth(now);

    const row = [
      formatDate(now),
      'Неделя ' + weekNum,
      getMonthName(now),
      now.getFullYear()
    ].concat(SOURCES.map(s => data[s] || 0)).concat([total]);

    sheet.appendRow(row);

    // Форматируем последнюю строку
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1, 1, row.length)
      .setBackground(lastRow % 2 === 0 ? '#f8f9fa' : '#ffffff');

    // Выделяем итого зелёным
    sheet.getRange(lastRow, row.length)
      .setBackground('#d4edda')
      .setFontWeight('bold');

    // Строим сводку для отправки
    let summary = '✅ *Данные сохранены!*\n\n';
    summary += '📅 ' + formatDate(now) + ', Неделя ' + weekNum + '\n';
    summary += '━━━━━━━━━━━━━━━━━━\n';
    SOURCES.forEach(s => {
      const val = data[s] || 0;
      const bar = '▓'.repeat(Math.min(val, 10)) + '░'.repeat(Math.max(0, 10 - Math.min(val, 10)));
      summary += (val > 0 ? '🔵' : '⚪') + ' ' + s + ': *' + val + '*\n';
    });
    summary += '━━━━━━━━━━━━━━━━━━\n';
    summary += '📊 *Итого за день: ' + total + '*\n\n';
    summary += '💾 Записано в таблицу "' + DATA_SHEET + '"';

    sendMessage(chatId, summary);
    props.deleteAllProperties();

  } catch (err) {
    Logger.log('saveToSheet error: ' + err.message);
    sendMessage(chatId,
      '❌ Ошибка записи в таблицу:\n' + err.message + '\n\n' +
      'Проверь SHEET_ID в настройках бота.'
    );
  }
}

// ============================================================
//  ЕЖЕДНЕВНЫЙ ТРИГГЕР — напоминание в 17:30
// ============================================================

function sendDailyReminder() {
  const now     = new Date();
  const day     = now.getDay(); // 0=вс, 6=сб
  if (day === 0 || day === 6) return; // не отправляем в выходные

  const weekNum = getWeekOfMonth(now);
  const today   = formatDate(now);

  const msg =
    '🔔 *Время вносить данные!*\n\n' +
    '📅 ' + today + ' — Неделя ' + weekNum + '\n\n' +
    'Нажми /начать чтобы внести данные\n' +
    'по источникам трафика за сегодня.\n\n' +
    '⏱ Займёт ~2 минуты';

  sendMessage(CHAT_ID, msg);
}

// ============================================================
//  ВСПОМОГАТЕЛЬНЫЕ КОМАНДЫ
// ============================================================

function showStatus(chatId) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(DATA_SHEET);

    if (!sheet || sheet.getLastRow() < 2) {
      sendMessage(chatId, '📭 Данных пока нет. Внеси первую запись через /начать');
      return;
    }

    const now       = new Date();
    const monthName = getMonthName(now);
    const data      = sheet.getDataRange().getValues();
    const headers   = data[0];

    // Фильтруем текущий месяц
    let monthTotal  = 0;
    let daysCount   = 0;
    const sourceTotals = {};
    SOURCES.forEach(s => sourceTotals[s] = 0);

    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === monthName && data[i][3] === now.getFullYear()) {
        daysCount++;
        SOURCES.forEach((s, idx) => {
          sourceTotals[s] += (data[i][4 + idx] || 0);
        });
        monthTotal += (data[i][4 + SOURCES.length] || 0);
      }
    }

    let msg = '📊 *Статистика за ' + monthName + '*\n';
    msg += '📅 Дней внесено: ' + daysCount + '\n';
    msg += '━━━━━━━━━━━━━━━━━━\n';
    SOURCES.forEach(s => {
      msg += '• ' + s + ': *' + sourceTotals[s] + '*\n';
    });
    msg += '━━━━━━━━━━━━━━━━━━\n';
    msg += '📈 *Итого за месяц: ' + monthTotal + '*';

    sendMessage(chatId, msg);

  } catch (err) {
    sendMessage(chatId, '❌ Ошибка: ' + err.message);
  }
}

function sendHelp(chatId) {
  sendMessage(chatId,
    '📖 *Команды бота:*\n\n' +
    '/начать — внести данные за сегодня\n' +
    '/статус — статистика за текущий месяц\n' +
    '/отмена — отменить текущий ввод\n' +
    '/помощь — это сообщение\n\n' +
    '⏰ Каждый будний день в 17:30\n' +
    'бот сам напомнит тебе внести данные.'
  );
}

// ============================================================
//  УТИЛИТЫ
// ============================================================

function sendMessage(chatId, text) {
  const url     = 'https://api.telegram.org/bot' + BOT_TOKEN + '/sendMessage';
  const payload = {
    chat_id    : String(chatId),
    text       : text,
    parse_mode : 'Markdown'
  };
  UrlFetchApp.fetch(url, {
    method      : 'post',
    contentType : 'application/json',
    payload     : JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

function formatDate(date) {
  const d = date.getDate().toString().padStart(2, '0');
  const m = (date.getMonth() + 1).toString().padStart(2, '0');
  const y = date.getFullYear();
  return d + '.' + m + '.' + y;
}

function getMonthName(date) {
  const months = [
    'Январь','Февраль','Март','Апрель','Май','Июнь',
    'Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'
  ];
  return months[date.getMonth()];
}

function getWeekOfMonth(date) {
  const firstDay = new Date(date.getFullYear(), date.getMonth(), 1).getDay();
  return Math.ceil((date.getDate() + (firstDay === 0 ? 6 : firstDay - 1)) / 7);
}

function getEmojiNum(n) {
  const emojis = ['','1️⃣','2️⃣','3️⃣','4️⃣','5️⃣','6️⃣','7️⃣'];
  return emojis[n] || n + '.';
}

// ============================================================
//  ПЕРВОНАЧАЛЬНАЯ НАСТРОЙКА — запускать один раз!
// ============================================================

/**
 * Шаг 1: Запусти эту функцию ОДИН РАЗ для подключения бота к таблице.
 * Сначала задеплой Apps Script как Web App (смотри инструкцию ниже).
 */
function setupWebhook() {
  const webAppUrl = ScriptApp.getService().getUrl();
  const apiUrl    = 'https://api.telegram.org/bot' + BOT_TOKEN +
                    '/setWebhook?url=' + encodeURIComponent(webAppUrl);

  const response  = UrlFetchApp.fetch(apiUrl);
  const result    = JSON.parse(response.getContentText());

  if (result.ok) {
    Logger.log('✅ Webhook установлен: ' + webAppUrl);
    Browser.msgBox('✅ Webhook установлен!\n\nURL: ' + webAppUrl);
  } else {
    Logger.log('❌ Ошибка: ' + JSON.stringify(result));
    Browser.msgBox('❌ Ошибка: ' + result.description);
  }
}

/**
 * Шаг 2: Запусти эту функцию ОДИН РАЗ для создания триггера 17:30.
 */
function setupDailyTrigger() {
  // Удаляем старые триггеры sendDailyReminder чтобы не дублировать
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendDailyReminder') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('sendDailyReminder')
    .timeBased()
    .everyDays(1)
    .atHour(17)
    .nearMinute(30)
    .create();

  Logger.log('✅ Триггер создан: каждый день в 17:30');
  Browser.msgBox('✅ Триггер создан!\nКаждый будний день в 17:30 бот будет присылать напоминание.');
}

/**
 * Проверка подключения — запусти для теста
 */
function testBot() {
  sendMessage(CHAT_ID,
    '🤖 Бот успешно подключён!\n\n' +
    'Напиши /начать чтобы внести данные.\n' +
    'Напиши /помощь для списка команд.'
  );
  Logger.log('Тестовое сообщение отправлено на chat_id: ' + CHAT_ID);
}

// ============================================================
//  ИНСТРУКЦИЯ ПО НАСТРОЙКЕ
// ============================================================
/*

ПОШАГОВАЯ ИНСТРУКЦИЯ:

── ШАГ 1: Создать Telegram бота ────────────────────────────
1. Открой Telegram → найди @BotFather
2. Напиши /newbot
3. Придумай название: «Маркетинг Отчёт»
4. Придумай username: например marketing_report_bot
5. BotFather пришлёт TOKEN — скопируй его в BOT_TOKEN выше

── ШАГ 2: Узнать свой Chat ID ──────────────────────────────
1. Найди в Telegram @userinfobot
2. Напиши /start
3. Бот пришлёт твой ID — вставь в CHAT_ID выше

── ШАГ 3: Получить ID таблицы ──────────────────────────────
1. Открой нужную Google таблицу
2. Скопируй из URL часть между /d/ и /edit
   Пример URL:
   https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms/edit
   ID = 1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms
3. Вставь в SHEET_ID выше

── ШАГ 4: Вставить код в Apps Script ───────────────────────
1. Открой Google таблицу
2. Расширения → Apps Script
3. Удали всё → вставь этот код
4. Сохрани (Ctrl+S)

── ШАГ 5: Задеплоить как Web App ───────────────────────────
1. Нажми «Развернуть» → «Новое развёртывание»
2. Тип: Web-приложение
3. Выполнять как: Я (your email)
4. Доступ: Все
5. Нажми «Развернуть»
6. Разреши доступ
7. Скопируй URL (он нужен для webhook)

── ШАГ 6: Запустить настройку ──────────────────────────────
1. В Apps Script выбери функцию setupWebhook → ▶ Выполнить
2. Затем выбери setupDailyTrigger → ▶ Выполнить
3. Затем выбери testBot → ▶ Выполнить
4. Открой Telegram — должно прийти сообщение от бота!

── ГОТОВО! ──────────────────────────────────────────────────
Каждый будний день в 17:30 бот пришлёт напоминание.
Напиши /начать → отвечай на 7 вопросов → данные запишутся в таблицу.

*/
