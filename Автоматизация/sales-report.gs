/**
 * ОТЧЁТ РОПа ДЛЯ ДИРЕКТОРА — Гидравлическое оборудование
 * Оптимизированная версия — минимум API-вызовов, работает за 15–20 сек
 *
 * КАК ИСПОЛЬЗОВАТЬ:
 * 1. Cmd+A → удалить старый код → вставить этот → Cmd+S
 * 2. Выбрать функцию createReport → нажать ▶ Выполнить
 */

// ─── КОНСТАНТЫ ───────────────────────────────
const MANAGERS = ['Бабурин И.', 'Волкова Т.', 'Микрюкова А.', 'Юргина И.', 'Ожегов А.', 'Лихачев А.'];
const PLAN = 10000000;

const C = {
  PURPLE_D: '#7B3F8C', PURPLE_L: '#E8D5F0',
  BLUE_D:   '#1F6AA5', BLUE_L:   '#CFE2F3',
  GREEN_D:  '#38761D', GREEN_L:  '#D9EAD3', GREEN_ROW: '#F0F7ED',
  ORANGE_D: '#E65100', ORANGE_L: '#FCE5CD', ORANGE_ROW: '#FEF0E7',
  DARK:     '#263238', LIGHT:    '#ECEFF1',
  WHITE: '#FFFFFF', BLACK: '#000000',
  OK: '#B7E1CD', WARN: '#FCE8B2', ERR: '#F4CCCC'
};

// ─── ГЛАВНАЯ ФУНКЦИЯ ─────────────────────────
function createReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const monthNames = ['Январь','Февраль','Март','Апрель','Май','Июнь',
                      'Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];
  const now = new Date();
  const sheetName = monthNames[now.getMonth()] + ' ' + now.getFullYear();

  // Удалить старый лист
  let old = ss.getSheetByName(sheetName);
  if (old) ss.deleteSheet(old);
  const sheet = ss.insertSheet(sheetName, 0);

  // Сводный лист
  buildSummarySheet(ss);

  // Строим все блоки
  buildPurple(sheet);
  buildBlue(sheet);
  buildTraffic(sheet);
  buildGreen(sheet);
  buildOrange(sheet);

  // Ширины столбцов одним батчем
  [140,110,110,110,130,110,110,120,120,120,120,110,110,110,110,110].forEach((w,i) => {
    sheet.setColumnWidth(i+1, w);
  });

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('✅ Отчёт «' + sheetName + '» создан!');
}

// ══════════════════════════════════════════════
// ФИОЛЕТОВАЯ (строки 1–5)
// ══════════════════════════════════════════════
function buildPurple(sheet) {
  // Заголовок
  sheet.getRange(1,1,1,8).merge().setValue('ПЛАН И ВЫПОЛНЕНИЕ')
    .setBackground(C.PURPLE_D).setFontColor(C.WHITE).setFontWeight('bold')
    .setFontSize(12).setHorizontalAlignment('center');

  // Строка 2 — подзаголовки
  sheet.getRange(2,1,1,8).setValues([[
    'Факт выручки (₽)','Дней прошло','% выполнения','План месяца (₽)',
    'Гидростанции','Гидроцилиндры','Прочее','План на день (₽)'
  ]]).setBackground(C.PURPLE_L).setFontWeight('bold').setFontSize(9);

  // Строка 3 — данные/формулы
  sheet.getRange(3,1).setValue(0).setNumberFormat('#,##0 ₽').setFontSize(11).setFontWeight('bold');
  sheet.getRange(3,2).setValue(0);
  sheet.getRange(3,3).setFormula('=IFERROR(A3/D3,0)').setNumberFormat('0.00%')
    .setFontSize(16).setFontWeight('bold').setBackground(C.PURPLE_L);
  sheet.getRange(3,4).setValue(PLAN).setNumberFormat('#,##0 ₽').setFontWeight('bold');
  sheet.getRange(3,5,1,3).setValues([[0,0,0]]);
  sheet.getRange(3,8).setFormula('=IFERROR(D3/30,0)').setNumberFormat('#,##0 ₽').setBackground(C.PURPLE_L);

  // Строка 4 — средний чек
  sheet.getRange(4,1).setValue('Средний чек (₽):').setBackground(C.PURPLE_L).setFontWeight('bold');
  sheet.getRange(4,2).setFormula('=IFERROR(A3/COUNTIF(A3,">"&0),0)')
    .setNumberFormat('#,##0.00 ₽').setBackground(C.PURPLE_L).setFontWeight('bold');

  // Фон белый для ввода
  sheet.getRange(3,1).setBackground(C.WHITE);
  sheet.getRange(3,2).setBackground(C.WHITE);
  sheet.getRange(3,4).setBackground(C.WHITE);
  sheet.getRange(3,5,1,3).setBackground(C.WHITE);

  // Условное форматирование % плана
  const r = sheet.getRange(3,3);
  const rules = sheet.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(1).setBackground(C.OK).setRanges([r]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.8,0.9999).setBackground(C.WARN).setRanges([r]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0.8).setBackground(C.ERR).setRanges([r]).build());
  sheet.setConditionalFormatRules(rules);

  sheet.getRange(1,1,4,8).setBorder(true,true,true,true,true,true,'#CCCCCC',SpreadsheetApp.BorderStyle.SOLID);
}

// ══════════════════════════════════════════════
// ГОЛУБАЯ (строки 6–14)
// ══════════════════════════════════════════════
function buildBlue(sheet) {
  const R = 6;

  sheet.getRange(R,1,1,8).merge().setValue('ЛИДОГЕНЕРАЦИЯ — ВХОДЯЩИЙ ТРАФИК ЗА МЕСЯЦ')
    .setBackground(C.BLUE_D).setFontColor(C.WHITE).setFontWeight('bold')
    .setFontSize(11).setHorizontalAlignment('center');

  sheet.getRange(R+1,1,1,7).setValues([[
    'Неделя','Новые лиды','Брак','Лиды в сделку','Старая база (звонки)','Отклики с базы','Конверсия лидов→сделки'
  ]]).setBackground(C.BLUE_L).setFontWeight('bold').setFontSize(9).setWrap(true);
  sheet.setRowHeight(R+1, 40);

  // 5 недель
  const weekData = [['Неделя 1',0,0,0,0,0],[' Неделя 2',0,0,0,0,0],['Неделя 3',0,0,0,0,0],['Неделя 4',0,0,0,0,0],['Неделя 5',0,0,0,0,0]];
  sheet.getRange(R+2,1,5,6).setValues(weekData).setBackground(C.WHITE);
  sheet.getRange(R+2,1,5,1).setBackground(C.BLUE_L).setFontWeight('bold');

  // Конверсия по неделям
  for (let i=0; i<5; i++) {
    const r = R+2+i;
    sheet.getRange(r,7).setFormula(`=IFERROR(D${r}/B${r},0)`).setNumberFormat('0.00%').setBackground(C.BLUE_L);
  }

  // Итого
  const tR = R+7;
  sheet.getRange(tR,1).setValue('ИТОГО').setBackground(C.BLUE_L).setFontWeight('bold');
  ['B','C','D','E','F'].forEach(col => {
    sheet.getRange(`${col}${tR}`).setFormula(`=SUM(${col}${R+2}:${col}${R+6})`).setBackground(C.BLUE_L).setFontWeight('bold');
  });
  sheet.getRange(tR,7).setFormula(`=IFERROR(D${tR}/B${tR},0)`)
    .setNumberFormat('0.00%').setBackground(C.BLUE_D).setFontColor(C.WHITE).setFontWeight('bold').setFontSize(13);

  sheet.getRange(R,1,tR-R+1,8).setBorder(true,true,true,true,true,true,'#CCCCCC',SpreadsheetApp.BorderStyle.SOLID);
}

// ══════════════════════════════════════════════
// РЕКЛАМНЫЕ ТРАФИКИ (строки 6–14, колонки 10–16)
// ══════════════════════════════════════════════
function buildTraffic(sheet) {
  const R = 6, COL = 10;

  sheet.getRange(R,COL,1,7).merge().setValue('РЕКЛАМНЫЕ ТРАФИКИ')
    .setBackground(C.BLUE_D).setFontColor(C.WHITE).setFontWeight('bold')
    .setFontSize(11).setHorizontalAlignment('center');

  sheet.getRange(R+1,COL,1,7).setValues([['Неделя','Яндекс Директ','SEO','Почта','Звонки','Carrotquest','Прочий трафик']])
    .setBackground(C.BLUE_L).setFontWeight('bold').setFontSize(9).setWrap(true);

  const wData = [['Неделя 1',0,0,0,0,0,0],['Неделя 2',0,0,0,0,0,0],['Неделя 3',0,0,0,0,0,0],['Неделя 4',0,0,0,0,0,0],['Неделя 5',0,0,0,0,0,0]];
  sheet.getRange(R+2,COL,5,7).setValues(wData).setBackground(C.WHITE);
  sheet.getRange(R+2,COL,5,1).setBackground(C.BLUE_L).setFontWeight('bold');

  // Итого
  const tR = R+7;
  sheet.getRange(tR,COL).setValue('ИТОГО').setBackground(C.BLUE_L).setFontWeight('bold');
  ['K','L','M','N','O','P'].forEach((col,i) => {
    sheet.getRange(tR,COL+1+i).setFormula(`=SUM(${col}${R+2}:${col}${R+6})`).setBackground(C.BLUE_L).setFontWeight('bold');
  });

  sheet.getRange(R,COL,tR-R+1,7).setBorder(true,true,true,true,true,true,'#CCCCCC',SpreadsheetApp.BorderStyle.SOLID);
  for (let c=COL; c<=COL+6; c++) sheet.setColumnWidth(c,110);
}

// ══════════════════════════════════════════════
// ЗЕЛЁНАЯ (строки 16–48)
// ══════════════════════════════════════════════
function buildGreen(sheet) {
  const R = 16;

  sheet.getRange(R,1,1,12).merge().setValue('ЗЕЛЁНАЯ ТАБЛИЦА — ЗАПОЛНЯТЬ ЕЖЕДНЕВНО!')
    .setBackground(C.GREEN_D).setFontColor(C.WHITE).setFontWeight('bold')
    .setFontSize(11).setHorizontalAlignment('center');

  sheet.getRange(R+1,1,1,12).setValues([[
    'Дата','Менеджер','Новые лиды','Брак','Новые сделки',
    'Старая база','Отклики со старой базы','Факт выручки (₽)',
    'Номер сделки','Наименование','Шт.','Постоплатные сделки'
  ]]).setBackground(C.GREEN_L).setFontWeight('bold').setFontSize(9).setWrap(true);
  sheet.setRowHeight(R+1, 40);

  // 30 строк — чередующийся фон
  for (let i=0; i<30; i++) {
    const r = R+2+i;
    const bg = i%2===0 ? C.GREEN_ROW : C.WHITE;
    sheet.getRange(r,1,1,12).setBackground(bg);
    sheet.getRange(r,1).setNumberFormat('dd.MM.yyyy');
    sheet.getRange(r,8).setNumberFormat('#,##0 ₽');
  }

  // Валидация — выпадающий список менеджеров
  sheet.getRange(R+2,2,30,1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(MANAGERS,true).setAllowInvalid(false).build()
  );

  sheet.getRange(R,1,32,12).setBorder(true,true,true,true,true,true,'#CCCCCC',SpreadsheetApp.BorderStyle.SOLID);
  sheet.setFrozenRows(R+1);
}

// ══════════════════════════════════════════════
// ОРАНЖЕВАЯ (строки 50–59)
// ══════════════════════════════════════════════
function buildOrange(sheet) {
  const R = 50;
  const GS = 18; // Green data start row
  const GE = 47; // Green data end row

  sheet.getRange(R,1,1,12).merge().setValue('СВОДКА ПО МЕНЕДЖЕРАМ (авто) — НЕ РЕДАКТИРОВАТЬ')
    .setBackground(C.ORANGE_D).setFontColor(C.WHITE).setFontWeight('bold')
    .setFontSize(11).setHorizontalAlignment('center');

  sheet.getRange(R+1,1,1,10).setValues([[
    'Менеджер','Новые лиды','Сделки','Отклики старая база','Старая база',
    'Конверсия лиды→сделки','Факт выручки (₽)','Конверсия в продажу','Постоплатные','Итог'
  ]]).setBackground(C.ORANGE_L).setFontWeight('bold').setFontSize(9).setWrap(true);
  sheet.setRowHeight(R+1, 45);

  MANAGERS.forEach((mgr, i) => {
    const r = R+2+i;
    const bg = i%2===0 ? C.ORANGE_ROW : C.WHITE;
    const bR = `B${GS}:B${GE}`;
    const q  = `"${mgr}"`;

    sheet.getRange(r,1).setValue(mgr).setBackground(C.ORANGE_L).setFontWeight('bold');
    sheet.getRange(r,2).setFormula(`=SUMIF(${bR},${q},C${GS}:C${GE})`).setBackground(bg);
    sheet.getRange(r,3).setFormula(`=SUMIF(${bR},${q},E${GS}:E${GE})`).setBackground(bg);
    sheet.getRange(r,4).setFormula(`=SUMIF(${bR},${q},G${GS}:G${GE})`).setBackground(bg);
    sheet.getRange(r,5).setFormula(`=SUMIF(${bR},${q},F${GS}:F${GE})`).setBackground(bg);
    sheet.getRange(r,6).setFormula(`=IFERROR(C${r}/B${r},0)`).setNumberFormat('0.00%').setBackground(bg);
    sheet.getRange(r,7).setFormula(`=SUMIF(${bR},${q},H${GS}:H${GE})`).setNumberFormat('#,##0 ₽').setBackground(bg);
    sheet.getRange(r,8).setFormula(`=IFERROR(C${r}/B${r},0)`).setNumberFormat('0.00%').setBackground(bg);
    sheet.getRange(r,9).setFormula(`=SUMIF(${bR},${q},L${GS}:L${GE})`).setBackground(bg);
    sheet.getRange(r,10).setFormula(`=IFERROR(G${r}/SUMIF(${bR},${q},H${GS}:H${GE})*100,0)`).setNumberFormat('0.00').setBackground(bg);
  });

  // Условное форматирование конверсии (колонка F)
  const convRng = sheet.getRange(R+2,6,MANAGERS.length,1);
  const rules = sheet.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.5).setBackground(C.OK).setRanges([convRng]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.2,0.4999).setBackground(C.WARN).setRanges([convRng]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0.2).setBackground(C.ERR).setRanges([convRng]).build());
  sheet.setConditionalFormatRules(rules);

  // Итого
  const tR = R+2+MANAGERS.length;
  sheet.getRange(tR,1).setValue('ИТОГО').setBackground(C.ORANGE_L).setFontWeight('bold');
  [2,3,4,5,9].forEach(c => {
    const col = String.fromCharCode(64+c);
    sheet.getRange(tR,c).setFormula(`=SUM(${col}${R+2}:${col}${R+1+MANAGERS.length})`).setBackground(C.ORANGE_L).setFontWeight('bold');
  });
  sheet.getRange(tR,6).setFormula(`=IFERROR(C${tR}/B${tR},0)`).setNumberFormat('0.00%').setBackground(C.ORANGE_L).setFontWeight('bold');
  sheet.getRange(tR,7).setFormula(`=SUM(G${R+2}:G${R+1+MANAGERS.length})`).setNumberFormat('#,##0 ₽').setBackground(C.ORANGE_L).setFontWeight('bold');

  sheet.getRange(R,1,tR-R+1,10).setBorder(true,true,true,true,true,true,'#CCCCCC',SpreadsheetApp.BorderStyle.SOLID);

  // Защита от случайного редактирования
  const prot = sheet.getRange(R,1,tR-R+1,10).protect();
  prot.setDescription('Оранжевая — только чтение').setWarningOnly(true);
}

// ══════════════════════════════════════════════
// СВОДНЫЙ ЛИСТ ЗА ГОД
// ══════════════════════════════════════════════
function buildSummarySheet(ss) {
  let s = ss.getSheetByName('Сводка за год');
  if (s) ss.deleteSheet(s);
  s = ss.insertSheet('Сводка за год');

  s.getRange(1,1,1,13).merge().setValue('СВОДНЫЙ ОТЧЁТ ПО ПРОДАЖАМ 2026')
    .setBackground(C.DARK).setFontColor(C.WHITE).setFontWeight('bold')
    .setFontSize(14).setHorizontalAlignment('center');

  s.getRange(2,1,1,13).setValues([[
    'Месяц','Факт выручки (₽)','% плана','Новых лидов','Сделок',
    'Конверсия','Гидростанции','Гидроцилиндры','Прочее',
    'Средний чек (₽)','Лучший менеджер','Примечания',''
  ]]).setBackground(C.DARK).setFontColor(C.WHITE).setFontWeight('bold').setFontSize(10).setWrap(true);
  s.setRowHeight(2,45);

  const months = ['Январь','Февраль','Март','Апрель','Май','Июнь',
                  'Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];
  months.forEach((m,i) => {
    const r = 3+i;
    const bg = i%2===0 ? C.LIGHT : C.WHITE;
    s.getRange(r,1).setValue(m+' 2026').setBackground(bg).setFontWeight('bold');
    s.getRange(r,2,1,11).setBackground(bg);
    s.getRange(r,2).setNumberFormat('#,##0 ₽');
    s.getRange(r,3).setNumberFormat('0.00%');
    s.getRange(r,6).setNumberFormat('0.00%');
    s.getRange(r,10).setNumberFormat('#,##0 ₽');
  });

  // Итого
  const tR = 15;
  s.getRange(tR,1,1,13).setBackground(C.DARK).setFontColor(C.WHITE);
  s.getRange(tR,1).setValue('ИТОГО ГОД').setFontWeight('bold').setFontSize(11);
  [2,4,5,7,8,9].forEach(c => {
    const col = String.fromCharCode(64+c);
    s.getRange(tR,c).setFormula(`=SUM(${col}3:${col}14)`).setFontWeight('bold').setFontColor(C.WHITE).setNumberFormat(c===2?'#,##0 ₽':'0');
  });
  s.getRange(tR,6).setFormula(`=IFERROR(E${tR}/D${tR},0)`).setNumberFormat('0.00%').setFontColor(C.WHITE);

  s.setColumnWidth(1,150);
  for (let c=2; c<=13; c++) s.setColumnWidth(c,120);

  s.getRange(1,1,tR,13).setBorder(true,true,true,true,true,true,'#CCCCCC',SpreadsheetApp.BorderStyle.SOLID);

  s.getRange(17,1,1,13).merge()
    .setValue('Заполнять вручную в конце каждого месяца или настроить автоимпорт через Битрикс24 → Make → Google Sheets')
    .setBackground('#FFF9C4').setWrap(true).setFontSize(9);
  s.setRowHeight(17,45);
}

// ══════════════════════════════════════════════
// АВТОЗАПУСК 1-го числа каждого месяца
// Запустить один раз вручную!
// ══════════════════════════════════════════════
function setupMonthlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('createReport').timeBased().onMonthDay(1).atHour(8).create();
  SpreadsheetApp.getUi().alert('✅ Триггер настроен: новая вкладка создаётся 1-го числа в 8:00');
}
