/**
 * ОТЧЁТ РОПа ДЛЯ ДИРЕКТОРА — Гидравлическое оборудование
 *
 * КАК ИСПОЛЬЗОВАТЬ:
 * 1. Открыть Google Таблицы (sheets.google.com) → создать новый файл
 * 2. Расширения → Apps Script → вставить этот код целиком
 * 3. Нажать ▶ Run (выбрать функцию createReport)
 * 4. Дать разрешения → таблица построится автоматически
 */

function createReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Получаем текущий месяц и год для названия вкладки
  const now = new Date();
  const monthNames = ['Январь','Февраль','Март','Апрель','Май','Июнь',
                      'Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];
  const sheetName = monthNames[now.getMonth()] + ' ' + now.getFullYear();

  // Удалить старый лист если есть
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(sheetName, 0);

  // Также создаём сводный лист
  createSummarySheet(ss);

  // ─────────────────────────────────────────────
  // БЛОК 1: ФИОЛЕТОВАЯ ТАБЛИЦА — ПЛАН/ФАКТ
  // ─────────────────────────────────────────────
  buildPurpleTable(sheet);

  // ─────────────────────────────────────────────
  // БЛОК 2: ГОЛУБАЯ ТАБЛИЦА — ЛИДОГЕНЕРАЦИЯ
  // ─────────────────────────────────────────────
  buildBlueTable(sheet);

  // ─────────────────────────────────────────────
  // БЛОК 3: РЕКЛАМНЫЕ ТРАФИКИ
  // ─────────────────────────────────────────────
  buildTrafficTable(sheet);

  // ─────────────────────────────────────────────
  // БЛОК 4: ЗЕЛЁНАЯ ТАБЛИЦА — ЕЖЕДНЕВНЫЙ ВВОД
  // ─────────────────────────────────────────────
  buildGreenTable(sheet);

  // ─────────────────────────────────────────────
  // БЛОК 5: ОРАНЖЕВАЯ ТАБЛИЦА — СВОДКА МЕНЕДЖЕРОВ
  // ─────────────────────────────────────────────
  buildOrangeTable(sheet);

  // Ширины столбцов
  sheet.setColumnWidth(1, 140);  // A
  sheet.setColumnWidth(2, 110);  // B
  sheet.setColumnWidth(3, 110);  // C
  sheet.setColumnWidth(4, 110);  // D
  sheet.setColumnWidth(5, 130);  // E
  sheet.setColumnWidth(6, 110);  // F
  sheet.setColumnWidth(7, 110);  // G
  sheet.setColumnWidth(8, 120);  // H
  sheet.setColumnWidth(9, 120);  // I
  sheet.setColumnWidth(10, 120); // J
  sheet.setColumnWidth(11, 120); // K
  sheet.setColumnWidth(12, 110); // L

  SpreadsheetApp.getUi().alert('✅ Отчёт «' + sheetName + '» создан успешно!');
}

// ══════════════════════════════════════════════
// ФИОЛЕТОВАЯ ТАБЛИЦА (строки 1–5)
// ══════════════════════════════════════════════
function buildPurpleTable(sheet) {
  const PURPLE_HEADER = '#7B3F8C';
  const PURPLE_LIGHT  = '#E8D5F0';
  const WHITE = '#FFFFFF';

  // Заголовок блока
  setCell(sheet, 1, 1, 'ПЛАН И ВЫПОЛНЕНИЕ', PURPLE_HEADER, WHITE, true, 12);
  sheet.getRange(1,1,1,8).merge().setHorizontalAlignment('center');

  // Строка 2: подзаголовки секций
  const headers2 = [
    ['Выполнение плана (выручка)', '', '% выполнения', 'План месяца (₽)',
     'Реализованная продукция', '', '', 'Финансы']
  ];
  sheet.getRange(2,1,1,8).setValues(headers2)
       .setBackground(PURPLE_LIGHT).setFontWeight('bold').setFontSize(9);

  // Строка 3: подзаголовки колонок
  const h3 = [['Факт выручки (₽)', 'Дней прошло', '=B4/D3*100&"%"',
               '', 'Гидростанции', 'Гидроцилиндры', 'Прочее',
               'План на день (₽)']];
  sheet.getRange(3,1,1,8).setValues(h3)
       .setBackground(PURPLE_LIGHT).setFontWeight('bold').setFontSize(8);

  // Строка 4: данные (вводятся РОПом, кроме формул)
  sheet.getRange(4,1).setBackground(WHITE).setFontSize(11).setFontWeight('bold')
       .setNumberFormat('#,##0 ₽');
  sheet.getRange(4,2).setValue(0).setBackground(WHITE); // дней прошло
  sheet.getRange(4,3).setFormula('=IFERROR(A4/D4,0)')
       .setNumberFormat('0.00%').setBackground(PURPLE_LIGHT).setFontWeight('bold').setFontSize(14);
  sheet.getRange(4,4).setValue(10000000).setBackground(WHITE)
       .setNumberFormat('#,##0 ₽').setFontWeight('bold');
  sheet.getRange(4,5).setValue(0).setBackground(WHITE); // гидростанции
  sheet.getRange(4,6).setValue(0).setBackground(WHITE); // гидроцилиндры
  sheet.getRange(4,7).setValue(0).setBackground(WHITE); // прочее
  sheet.getRange(4,8).setFormula('=IFERROR(D4/30,0)')
       .setNumberFormat('#,##0 ₽').setBackground(PURPLE_LIGHT);

  // Строка 5: средний чек
  setCell(sheet, 5, 1, 'Средний чек (₽):', PURPLE_LIGHT, '#000000', true, 9);
  sheet.getRange(5,2).setFormula('=IFERROR(A4/SUMIF(A4:A4,">"&0),0)')
       .setNumberFormat('#,##0.00 ₽').setBackground(PURPLE_LIGHT).setFontWeight('bold');

  // Условное форматирование % выполнения
  const planCell = sheet.getRange(4,3);
  const rules = sheet.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(1)
    .setBackground('#B7E1CD').setRanges([planCell]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0.8, 0.9999)
    .setBackground('#FCE8B2').setRanges([planCell]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0.8)
    .setBackground('#F4CCCC').setRanges([planCell]).build());
  sheet.setConditionalFormatRules(rules);

  addBorder(sheet.getRange(1,1,5,8));
}

// ══════════════════════════════════════════════
// ГОЛУБАЯ ТАБЛИЦА (строки 7–14)
// ══════════════════════════════════════════════
function buildBlueTable(sheet) {
  const BLUE_HEADER = '#1F6AA5';
  const BLUE_LIGHT  = '#CFE2F3';
  const WHITE = '#FFFFFF';
  const ROW = 7;

  setCell(sheet, ROW, 1, 'ЛИДОГЕНЕРАЦИЯ — ВХОДЯЩИЙ ТРАФИК ЗА МЕСЯЦ', BLUE_HEADER, WHITE, true, 11);
  sheet.getRange(ROW,1,1,8).merge().setHorizontalAlignment('center');

  const h = [['Неделя', 'Новые лиды', 'Брак', 'Лиды в сделку',
              'Старая база (звонки)', 'Отклики с базы', 'Конверсия лидов в сделку', '']];
  sheet.getRange(ROW+1,1,1,8).setValues(h)
       .setBackground(BLUE_LIGHT).setFontWeight('bold').setFontSize(9)
       .setWrap(true).setVerticalAlignment('middle');
  sheet.setRowHeight(ROW+1, 40);

  const weeks = ['Неделя 1','Неделя 2','Неделя 3','Неделя 4','Неделя 5'];
  for (let i = 0; i < 5; i++) {
    const r = ROW + 2 + i;
    sheet.getRange(r,1).setValue(weeks[i]).setBackground(BLUE_LIGHT).setFontWeight('bold');
    sheet.getRange(r,2,1,6).setBackground(WHITE);
    // Конверсия = лиды в сделку / новые лиды
    const col = columnLetter(r, 4); // D
    const col2 = columnLetter(r, 2); // B
    sheet.getRange(r,7).setFormula(`=IFERROR(D${r}/B${r},0)`)
         .setNumberFormat('0.00%').setBackground(BLUE_LIGHT);
  }

  // Итого
  const totalRow = ROW + 7;
  setCell(sheet, totalRow, 1, 'ИТОГО', BLUE_LIGHT, '#000000', true, 10);
  for (let c = 2; c <= 6; c++) {
    const startR = ROW + 2;
    const endR   = ROW + 6;
    const col = String.fromCharCode(64 + c);
    sheet.getRange(totalRow, c).setFormula(`=SUM(${col}${startR}:${col}${endR})`)
         .setBackground(BLUE_LIGHT).setFontWeight('bold');
  }
  sheet.getRange(totalRow, 7).setFormula(`=IFERROR(D${totalRow}/B${totalRow},0)`)
       .setNumberFormat('0.00%').setBackground('#1F6AA5').setFontColor(WHITE).setFontWeight('bold').setFontSize(12);

  addBorder(sheet.getRange(ROW, 1, totalRow - ROW + 1, 8));
}

// ══════════════════════════════════════════════
// РЕКЛАМНЫЕ ТРАФИКИ (строки 7–14, колонки 10–16)
// ══════════════════════════════════════════════
function buildTrafficTable(sheet) {
  const BLUE_HEADER = '#1F6AA5';
  const BLUE_LIGHT  = '#CFE2F3';
  const WHITE = '#FFFFFF';
  const ROW = 7;
  const COL = 10; // J

  setCell(sheet, ROW, COL, 'РЕКЛАМНЫЕ ТРАФИКИ', BLUE_HEADER, WHITE, true, 11);
  sheet.getRange(ROW, COL, 1, 7).merge().setHorizontalAlignment('center');

  const h = [['Неделя', 'Яндекс Директ', 'SEO', 'Почта', 'Звонки', 'Carrotquest', 'Прочий трафик']];
  sheet.getRange(ROW+1, COL, 1, 7).setValues(h)
       .setBackground(BLUE_LIGHT).setFontWeight('bold').setFontSize(9)
       .setWrap(true).setVerticalAlignment('middle');

  const weeks = ['Неделя 1','Неделя 2','Неделя 3','Неделя 4','Неделя 5'];
  for (let i = 0; i < 5; i++) {
    const r = ROW + 2 + i;
    sheet.getRange(r, COL).setValue(weeks[i]).setBackground(BLUE_LIGHT).setFontWeight('bold');
    sheet.getRange(r, COL+1, 1, 6).setBackground(WHITE);
  }

  const totalRow = ROW + 7;
  setCell(sheet, totalRow, COL, 'ИТОГО', BLUE_LIGHT, '#000000', true, 10);
  const cols = ['K','L','M','N','O','P'];
  cols.forEach((col, idx) => {
    const startR = ROW + 2;
    const endR   = ROW + 6;
    sheet.getRange(totalRow, COL + 1 + idx)
         .setFormula(`=SUM(${col}${startR}:${col}${endR})`)
         .setBackground(BLUE_LIGHT).setFontWeight('bold');
  });

  addBorder(sheet.getRange(ROW, COL, totalRow - ROW + 1, 7));

  // Ширины доп. столбцов
  for (let c = 10; c <= 16; c++) sheet.setColumnWidth(c, 110);
}

// ══════════════════════════════════════════════
// ЗЕЛЁНАЯ ТАБЛИЦА (строки 16–50)
// ══════════════════════════════════════════════
function buildGreenTable(sheet) {
  const GREEN_HEADER = '#38761D';
  const GREEN_LIGHT  = '#D9EAD3';
  const WHITE = '#FFFFFF';
  const ROW = 16;

  setCell(sheet, ROW, 1, '⚠ ЗЕЛЁНАЯ ТАБЛИЦА — ЗАПОЛНЯТЬ ЕЖЕДНЕВНО!', GREEN_HEADER, '#FFFFFF', true, 11);
  sheet.getRange(ROW, 1, 1, 12).merge().setHorizontalAlignment('center');

  const h = [['Дата', 'Менеджер', 'Новые лиды', 'Брак', 'Новые сделки',
              'Старая база', 'Отклики со старой базы', 'Факт выручки (₽)',
              'Номер сделки', 'Наименование', 'Шт.', 'Постоплатные сделки']];
  sheet.getRange(ROW+1, 1, 1, 12).setValues(h)
       .setBackground(GREEN_LIGHT).setFontWeight('bold').setFontSize(9)
       .setWrap(true).setVerticalAlignment('middle');
  sheet.setRowHeight(ROW+1, 40);

  // 30 строк для ввода
  const managers = ['Бабурин И.', 'Волкова Т.', 'Микрюкова А.', 'Юргина И.', 'Ожегов А.', 'Лихачев А.'];
  for (let i = 0; i < 30; i++) {
    const r = ROW + 2 + i;
    const bg = i % 2 === 0 ? '#F0F7ED' : WHITE;
    sheet.getRange(r, 1).setBackground(bg).setNumberFormat('dd.MM.yyyy');
    sheet.getRange(r, 2).setBackground(bg); // менеджер
    sheet.getRange(r, 3, 1, 10).setBackground(bg);
    sheet.getRange(r, 8).setNumberFormat('#,##0 ₽');
  }

  // Валидация менеджеров
  const managerRange = sheet.getRange(ROW+2, 2, 30, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(managers, true)
    .setAllowInvalid(false).build();
  managerRange.setDataValidation(rule);

  addBorder(sheet.getRange(ROW, 1, 32, 12));

  // Закрепить шапку
  sheet.setFrozenRows(ROW + 1);
}

// ══════════════════════════════════════════════
// ОРАНЖЕВАЯ ТАБЛИЦА (строки 52–60)
// ══════════════════════════════════════════════
function buildOrangeTable(sheet) {
  const ORANGE_HEADER = '#E65100';
  const ORANGE_LIGHT  = '#FCE5CD';
  const WHITE = '#FFFFFF';
  const ROW = 52;
  const GREEN_DATA_START = 18; // строка где начинаются данные зелёной таблицы
  const GREEN_DATA_END   = 47; // строка где заканчиваются

  setCell(sheet, ROW, 1, 'СВОДКА ПО МЕНЕДЖЕРАМ (авто из зелёной таблицы) — НЕ РЕДАКТИРОВАТЬ', ORANGE_HEADER, '#FFFFFF', true, 11);
  sheet.getRange(ROW, 1, 1, 12).merge().setHorizontalAlignment('center');

  const h = [['Менеджер', 'Новые лиды', 'Сделки', 'Отклики старая база',
              'Старая база', 'Конверсия лиды→сделки', 'Факт выручки (₽)',
              'Конверсия в продажу', 'Постоплатные сделки', 'Реализованные сделки', '', '']];
  sheet.getRange(ROW+1, 1, 1, 12).setValues(h)
       .setBackground(ORANGE_LIGHT).setFontWeight('bold').setFontSize(9)
       .setWrap(true).setVerticalAlignment('middle');
  sheet.setRowHeight(ROW+1, 50);

  const managers = ['Бабурин И.', 'Волкова Т.', 'Микрюкова А.', 'Юргина И.', 'Ожегов А.', 'Лихачев А.'];

  managers.forEach((mgr, i) => {
    const r = ROW + 2 + i;
    sheet.getRange(r, 1).setValue(mgr).setBackground(ORANGE_LIGHT).setFontWeight('bold');

    // Формулы SUMIF по менеджеру из зелёной таблицы
    // Зелёная: B=менеджер, C=новые лиды, D=брак, E=новые сделки, F=старая база, G=отклики, H=выручка, L=постоплата
    const mgrRef = `"${mgr}"`;
    const bRange = `B${GREEN_DATA_START}:B${GREEN_DATA_END}`;

    sheet.getRange(r, 2).setFormula(`=SUMIF(${bRange},${mgrRef},C${GREEN_DATA_START}:C${GREEN_DATA_END})`); // лиды
    sheet.getRange(r, 3).setFormula(`=SUMIF(${bRange},${mgrRef},E${GREEN_DATA_START}:E${GREEN_DATA_END})`); // сделки
    sheet.getRange(r, 4).setFormula(`=SUMIF(${bRange},${mgrRef},G${GREEN_DATA_START}:G${GREEN_DATA_END})`); // отклики
    sheet.getRange(r, 5).setFormula(`=SUMIF(${bRange},${mgrRef},F${GREEN_DATA_START}:F${GREEN_DATA_END})`); // старая база
    // Конверсия лиды→сделки (IFERROR убирает #DIV/0!)
    sheet.getRange(r, 6).setFormula(`=IFERROR(C${r}/B${r},0)`).setNumberFormat('0.00%');
    sheet.getRange(r, 7).setFormula(`=SUMIF(${bRange},${mgrRef},H${GREEN_DATA_START}:H${GREEN_DATA_END})`).setNumberFormat('#,##0 ₽');
    // Конверсия в продажу
    sheet.getRange(r, 8).setFormula(`=IFERROR(C${r}/B${r},0)`).setNumberFormat('0.00%');
    sheet.getRange(r, 9).setFormula(`=SUMIF(${bRange},${mgrRef},L${GREEN_DATA_START}:L${GREEN_DATA_END})`); // постоплата
    sheet.getRange(r, 10).setFormula(`=IFERROR(C${r}/B${r},0)`).setNumberFormat('0.00%');

    const bg = i % 2 === 0 ? '#FEF0E7' : WHITE;
    sheet.getRange(r, 2, 1, 9).setBackground(bg);

    // Условное форматирование конверсии
    const convRange = sheet.getRange(r, 6);
    const rules = sheet.getConditionalFormatRules();
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(0.5).setBackground('#B7E1CD').setRanges([convRange]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(0.2, 0.4999).setBackground('#FCE8B2').setRanges([convRange]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0.2).setBackground('#F4CCCC').setRanges([convRange]).build());
    sheet.setConditionalFormatRules(rules);
  });

  // Итого (6 менеджеров: строки ROW+2 … ROW+7)
  const totalRow = ROW + 8;
  setCell(sheet, totalRow, 1, 'ИТОГО', ORANGE_LIGHT, '#000000', true, 10);
  const sumCols = [2, 3, 4, 5, 7, 9];
  sumCols.forEach(c => {
    const col = String.fromCharCode(64 + c);
    sheet.getRange(totalRow, c).setFormula(`=SUM(${col}${ROW+2}:${col}${ROW+7})`)
         .setBackground(ORANGE_LIGHT).setFontWeight('bold');
    if (c === 7) sheet.getRange(totalRow, c).setNumberFormat('#,##0 ₽');
  });
  sheet.getRange(totalRow, 6)
       .setFormula(`=IFERROR(C${totalRow}/B${totalRow},0)`)
       .setNumberFormat('0.00%').setBackground(ORANGE_LIGHT).setFontWeight('bold');

  addBorder(sheet.getRange(ROW, 1, totalRow - ROW + 1, 10));

  // Защита оранжевой таблицы
  const protection = sheet.getRange(ROW, 1, totalRow - ROW + 1, 10).protect();
  protection.setDescription('Оранжевая таблица — только чтение');
  protection.setWarningOnly(true); // предупреждение, но не блокировка
}

// ══════════════════════════════════════════════
// СВОДНЫЙ ЛИСТ ЗА ГОД
// ══════════════════════════════════════════════
function createSummarySheet(ss) {
  let summary = ss.getSheetByName('📊 Сводка за год');
  if (summary) ss.deleteSheet(summary);
  summary = ss.insertSheet('📊 Сводка за год');

  const DARK = '#263238';
  const LIGHT = '#ECEFF1';
  const WHITE = '#FFFFFF';

  setCell(summary, 1, 1, 'СВОДНЫЙ ОТЧЁТ ПО ПРОДАЖАМ — ГОД', DARK, WHITE, true, 14);
  summary.getRange(1,1,1,14).merge().setHorizontalAlignment('center');

  const h = [['Месяц', 'Факт выручки (₽)', '% плана', 'Новых лидов', 'Сделок',
              'Конверсия', 'Гидростанции', 'Гидроцилиндры', 'Прочее',
              'Средний чек (₽)', 'Менеджеров активных', 'Лучший менеджер', 'Примечания', '']];
  summary.getRange(2,1,1,14).setValues(h)
         .setBackground(DARK).setFontColor(WHITE).setFontWeight('bold').setFontSize(10)
         .setWrap(true).setVerticalAlignment('middle');
  summary.setRowHeight(2, 45);

  const months = ['Январь','Февраль','Март','Апрель','Май','Июнь',
                  'Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];
  months.forEach((m, i) => {
    const r = 3 + i;
    const bg = i % 2 === 0 ? LIGHT : WHITE;
    summary.getRange(r, 1).setValue(m + ' 2026').setBackground(bg).setFontWeight('bold');
    summary.getRange(r, 2, 1, 12).setBackground(bg);
    summary.getRange(r, 2).setNumberFormat('#,##0 ₽');
    summary.getRange(r, 3).setNumberFormat('0.00%');
    summary.getRange(r, 6).setNumberFormat('0.00%');
    summary.getRange(r, 10).setNumberFormat('#,##0 ₽');
  });

  // Итого год
  const totalRow = 15;
  setCell(summary, totalRow, 1, 'ИТОГО ГОД', DARK, WHITE, true, 11);
  summary.getRange(totalRow, 1, 1, 13).setBackground(DARK).setFontColor(WHITE);
  [2, 4, 5, 7, 8, 9].forEach(c => {
    const col = String.fromCharCode(64 + c);
    summary.getRange(totalRow, c).setFormula(`=SUM(${col}3:${col}14)`)
           .setFontWeight('bold').setFontColor(WHITE).setBackground(DARK);
    if (c === 2) summary.getRange(totalRow, c).setNumberFormat('#,##0 ₽');
  });
  summary.getRange(totalRow, 6).setFormula(`=IFERROR(E${totalRow}/D${totalRow},0)`)
         .setNumberFormat('0.00%').setFontColor(WHITE).setBackground(DARK);

  // Ширины
  summary.setColumnWidth(1, 150);
  for (let c = 2; c <= 13; c++) summary.setColumnWidth(c, 130);

  addBorder(summary.getRange(1, 1, totalRow, 13));

  // Примечание по заполнению
  setCell(summary, 17, 1, '💡 Как заполнять: данные по каждому месяцу вносите вручную из месячных вкладок, либо настройте автоимпорт через Битрикс24 → Make → Google Sheets', '#FFF9C4', '#000000', false, 9);
  summary.getRange(17, 1, 1, 13).merge().setWrap(true);
  summary.setRowHeight(17, 50);
}

// ══════════════════════════════════════════════
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ══════════════════════════════════════════════
function setCell(sheet, row, col, value, bgColor, fontColor, bold, fontSize) {
  const cell = sheet.getRange(row, col);
  cell.setValue(value)
      .setBackground(bgColor)
      .setFontColor(fontColor)
      .setFontWeight(bold ? 'bold' : 'normal')
      .setFontSize(fontSize);
}

function addBorder(range) {
  range.setBorder(true, true, true, true, true, true,
    '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
}

function columnLetter(row, col) {
  return String.fromCharCode(64 + col) + row;
}

// ══════════════════════════════════════════════
// ФУНКЦИЯ ДЛЯ СОЗДАНИЯ НОВОЙ ВКЛАДКИ МЕСЯЦА
// (запускается в начале каждого месяца)
// ══════════════════════════════════════════════
function createNewMonth() {
  createReport();
}

// ══════════════════════════════════════════════
// НАСТРОЙКА АВТОЗАПУСКА — создаёт триггер на 1-е число каждого месяца
// Запустить один раз вручную!
// ══════════════════════════════════════════════
function setupMonthlyTrigger() {
  // Удалить старые триггеры
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // Создать новый — каждый месяц 1-го числа
  ScriptApp.newTrigger('createNewMonth')
    .timeBased()
    .onMonthDay(1)
    .atHour(8)
    .create();

  SpreadsheetApp.getUi().alert('✅ Триггер настроен: новая вкладка будет создаваться автоматически 1-го числа каждого месяца в 8:00');
}
