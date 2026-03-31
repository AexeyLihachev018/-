// ============================================================
//  TELEGRAM BOT — Отчёт РК Маркетинг
//  Режим: POLLING (проверка сообщений каждую минуту)
//  Не требует webhook и развёртывания!
// ============================================================

const BOT_TOKEN  = 'YOUR_BOT_TOKEN';       // получи у @BotFather в Telegram
const CHAT_ID    = 'YOUR_CHAT_ID';         // узнай через @userinfobot
const SHEET_ID   = 'YOUR_GOOGLE_SHEET_ID'; // ID таблицы из URL
const DATA_SHEET = 'Данные бота';

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
//  ПЕРВЫЙ ЗАПУСК — выполни ОДИН РАЗ
// ============================================================

function setupBot() {
  // 1. Удаляем webhook (чтобы не мешал polling)
  UrlFetchApp.fetch(
    'https://api.telegram.org/bot' + BOT_TOKEN + '/deleteWebhook?drop_pending_updates=true',
    { method: 'post', muteHttpExceptions: true }
  );

  // 2. Удаляем старые триггеры
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // 3. Триггер polling — каждую минуту
  ScriptApp.newTrigger('pollUpdates').timeBased().everyMinutes(1).create();

  // 4. Триггер напоминания — 17:30 каждый день
  ScriptApp.newTrigger('sendDailyReminder').timeBased().everyDays(1).atHour(17).nearMinute(30).create();

  // 5. Сбрасываем счётчик обновлений
  PropertiesService.getScriptProperties().deleteAllProperties();

  // 6. Тестовое сообщение
  sendMessage(CHAT_ID, '🤖 Бот запущен!\n\nНапиши /начать чтобы внести данные.\nНапиши /помощь для списка команд.');

  Browser.msgBox('✅ Бот настроен!\n\nТриггеры созданы:\n• Polling: каждую минуту\n• Напоминание: 17:30\n\nПроверь Telegram — должно прийти сообщение!');
}

// ============================================================
//  POLLING — проверяет новые сообщения каждую минуту
// ============================================================

function pollUpdates() {
  const props       = PropertiesService.getScriptProperties();
  const lastId      = parseInt(props.getProperty('last_update_id') || '0');
  const url         = 'https://api.telegram.org/bot' + BOT_TOKEN +
                      '/getUpdates?offset=' + (lastId + 1) + '&limit=10&timeout=0';

  const response    = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const data        = JSON.parse(response.getContentText());

  if (!data.ok || !data.result || data.result.length === 0) return;

  data.result.forEach(update => {
    if (update.message) {
      const chatId = String(update.message.chat.id);
      const text   = (update.message.text || '').trim();

      if (chatId === String(CHAT_ID)) {
        handleMessage(chatId, text);
      }
    }
    props.setProperty('last_update_id', String(update.update_id));
  });
}

// ============================================================
//  ЛОГИКА ДИАЛОГА
// ============================================================

function handleMessage(chatId, text) {
  const props = PropertiesService.getScriptProperties();
  const state = parseInt(props.getProperty('dialog_state') || '0');

  if (text === '/start' || text === '/начать') {
    startDialog(chatId, props);
    return;
  }
  if (text === '/отмена' || text === '/cancel') {
    props.deleteProperty('dialog_state');
    props.deleteProperty('dialog_data');
    sendMessage(chatId, '❌ Ввод отменён. Напиши /начать чтобы начать заново.');
    return;
  }
  if (text === '/статус') { showStatus(chatId); return; }
  if (text === '/помощь' || text === '/help') { sendHelp(chatId); return; }

  if (state >= 1 && state <= SOURCES.length) {
    processDataInput(chatId, text, state, props);
    return;
  }

  sendMessage(chatId,
    '👋 Привет!\n\nНапиши /начать чтобы внести данные за сегодня.\nИли /помощь для списка команд.'
  );
}

function startDialog(chatId, props) {
  props.setProperty('dialog_state', '1');
  props.setProperty('dialog_data', JSON.stringify({}));

  sendMessage(chatId,
    '📊 *Отчёт РК Маркетинг*\n' +
    '📅 ' + formatDate(new Date()) + ' (Неделя ' + getWeekOfMonth(new Date()) + ')\n\n' +
    'Вводи только цифры (количество обращений).\nДля отмены: /отмена\n\n' +
    '━━━━━━━━━━━━━━━━━━\n' +
    '1️⃣ *' + SOURCES[0] + '*\nСколько обращений пришло сегодня?'
  );
}

function processDataInput(chatId, text, state, props) {
  const value = parseInt(text);
  if (isNaN(value) || value < 0) {
    sendMessage(chatId,
      '⚠️ Введи только число (например: 5 или 0)\n\n' +
      getEmojiNum(state) + ' *' + SOURCES[state - 1] + '*\nСколько обращений?'
    );
    return;
  }

  const data = JSON.parse(props.getProperty('dialog_data') || '{}');
  data[SOURCES[state - 1]] = value;
  props.setProperty('dialog_data', JSON.stringify(data));

  if (state < SOURCES.length) {
    const next = state + 1;
    props.setProperty('dialog_state', String(next));
    sendMessage(chatId,
      '✅ ' + SOURCES[state - 1] + ' = *' + value + '*\n\n' +
      '━━━━━━━━━━━━━━━━━━\n' +
      getEmojiNum(next) + ' *' + SOURCES[next - 1] + '*\nСколько обращений?'
    );
  } else {
    props.deleteProperty('dialog_state');
    data[SOURCES[state - 1]] = value;
    saveToSheet(chatId, data, props);
  }
}

function saveToSheet(chatId, data) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    let sheet   = ss.getSheetByName(DATA_SHEET);

    if (!sheet) {
      sheet = ss.insertSheet(DATA_SHEET);
      const headers = ['Дата', 'Неделя', 'Месяц', 'Год'].concat(SOURCES).concat(['ИТОГО']);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers])
           .setBackground('#4a4a4a').setFontColor('#ffffff').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    const now     = new Date();
    const total   = SOURCES.reduce((sum, s) => sum + (data[s] || 0), 0);
    const weekNum = getWeekOfMonth(now);
    const row     = [formatDate(now), 'Неделя ' + weekNum, getMonthName(now), now.getFullYear()]
                    .concat(SOURCES.map(s => data[s] || 0)).concat([total]);

    sheet.appendRow(row);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1, 1, row.length).setBackground(lastRow % 2 === 0 ? '#f8f9fa' : '#ffffff');
    sheet.getRange(lastRow, row.length).setBackground('#d4edda').setFontWeight('bold');

    let summary = '✅ *Данные сохранены!*\n\n📅 ' + formatDate(now) + ', Неделя ' + weekNum + '\n━━━━━━━━━━━━━━━━━━\n';
    SOURCES.forEach(s => {
      summary += (data[s] > 0 ? '🔵' : '⚪') + ' ' + s + ': *' + (data[s] || 0) + '*\n';
    });
    summary += '━━━━━━━━━━━━━━━━━━\n📊 *Итого за день: ' + total + '*\n\n💾 Записано в таблицу';

    sendMessage(chatId, summary);
    PropertiesService.getScriptProperties().deleteProperty('dialog_data');

  } catch (err) {
    Logger.log('saveToSheet error: ' + err.message);
    sendMessage(chatId, '❌ Ошибка записи в таблицу:\n' + err.message);
  }
}

// ============================================================
//  ЕЖЕДНЕВНОЕ НАПОМИНАНИЕ — 17:30
// ============================================================

function sendDailyReminder() {
  const day = new Date().getDay();
  if (day === 0 || day === 6) return;
  sendMessage(CHAT_ID,
    '🔔 *Время вносить данные!*\n\n' +
    '📅 ' + formatDate(new Date()) + ' — Неделя ' + getWeekOfMonth(new Date()) + '\n\n' +
    'Нажми /начать чтобы внести данные по источникам трафика за сегодня.\n\n⏱ Займёт ~2 минуты'
  );
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
    const now        = new Date();
    const monthName  = getMonthName(now);
    const rows       = sheet.getDataRange().getValues();
    let monthTotal   = 0, daysCount = 0;
    const totals     = {};
    SOURCES.forEach(s => totals[s] = 0);

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][2] === monthName && rows[i][3] === now.getFullYear()) {
        daysCount++;
        SOURCES.forEach((s, idx) => { totals[s] += (rows[i][4 + idx] || 0); });
        monthTotal += (rows[i][4 + SOURCES.length] || 0);
      }
    }
    let msg = '📊 *Статистика за ' + monthName + '*\n📅 Дней: ' + daysCount + '\n━━━━━━━━━━━━━━━━━━\n';
    SOURCES.forEach(s => { msg += '• ' + s + ': *' + totals[s] + '*\n'; });
    msg += '━━━━━━━━━━━━━━━━━━\n📈 *Итого: ' + monthTotal + '*';
    sendMessage(chatId, msg);
  } catch (err) {
    sendMessage(chatId, '❌ Ошибка: ' + err.message);
  }
}

function sendHelp(chatId) {
  sendMessage(chatId,
    '📖 *Команды бота:*\n\n' +
    '/начать — внести данные за сегодня\n' +
    '/статус — статистика за месяц\n' +
    '/отмена — отменить текущий ввод\n' +
    '/помощь — это сообщение\n\n' +
    '⏰ Каждый будний день в 17:30 — напоминание.'
  );
}

// ============================================================
//  УТИЛИТЫ
// ============================================================

function sendMessage(chatId, text) {
  UrlFetchApp.fetch('https://api.telegram.org/bot' + BOT_TOKEN + '/sendMessage', {
    method      : 'post',
    contentType : 'application/json',
    payload     : JSON.stringify({ chat_id: String(chatId), text: text, parse_mode: 'Markdown' }),
    muteHttpExceptions: true
  });
}

function formatDate(d) {
  return d.getDate().toString().padStart(2,'0') + '.' +
         (d.getMonth()+1).toString().padStart(2,'0') + '.' + d.getFullYear();
}

function getMonthName(d) {
  return ['Январь','Февраль','Март','Апрель','Май','Июнь',
          'Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'][d.getMonth()];
}

function getWeekOfMonth(d) {
  const firstDay = new Date(d.getFullYear(), d.getMonth(), 1).getDay();
  return Math.ceil((d.getDate() + (firstDay === 0 ? 6 : firstDay - 1)) / 7);
}

function getEmojiNum(n) {
  return (['','1️⃣','2️⃣','3️⃣','4️⃣','5️⃣','6️⃣','7️⃣'])[n] || n + '.';
}
