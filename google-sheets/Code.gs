// Связка бота Карины с Google Таблицей.
//
// Лист расписания на месяц (например «Октябрь 2026»): каждая строка — слот 30 минут.
// Столбцы: Дата | День недели | Время | ФИО | Процедура | ИИН | Телефон | Статус слота
//
// POST (doPost) — бот присылает JSON из блока <crm>...</crm>:
//   • статус «Записан» или «Перенёс запись» — ФИО, процедура, ИИН и телефон
//     записываются в слот «Дата визита» на листе расписания;
//   • «Отменил запись» — слот освобождается;
//   • любой статус дополнительно пишется строкой в лист «Клиенты» (журнал).
// GET (doGet) — свободные слоты для бота: ?secret=...&days=7
//
// createCurrentMonth / createNextMonth — создать лист расписания на месяц
// (запускать из меню «Расписание» в таблице).

const SECRET = 'ЗАМЕНИТЕ_НА_СВОЙ_СЕКРЕТНЫЙ_КЛЮЧ';
const LOG_SHEET = 'Клиенты';
const LOG_HEADERS = ['Дата и время', 'ФИО', 'ИИН', 'Телефон', 'Дата визита', 'Напомнить', 'Статус'];
const SCHEDULE_HEADERS = ['Дата', 'День недели', 'Время', 'ФИО', 'Процедура', 'ИИН', 'Телефон', 'Статус слота'];
const PRIMARY = 'Первичная диагностика ЖКТ';
const DAYS = ['Воскресенье', 'Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота'];
const MONTHS = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль',
  'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
const TZ = 'Asia/Almaty';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Расписание')
    .addItem('Создать лист на текущий месяц', 'createCurrentMonth')
    .addItem('Создать лист на следующий месяц', 'createNextMonth')
    .addToUi();
}

// ---------- Веб-приложение ----------

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  if (data.secret !== SECRET) return json({ ok: false, error: 'forbidden' });

  const iin = String(data['ИИН'] || '').replace(/\D/g, '');
  const rec = {
    fio: data['ФИО'] || '',
    iin: iin.length === 12 ? iin : '',
    phone: data['Телефон'] || '',
    visit: data['Дата визита'] || '',
    remind: data['Напомнить'] || '',
    status: data['Статус'] || '',
  };

  logRow(rec);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (rec.status === 'Записан' || rec.status === 'Перенёс запись') {
      clearSlotsOf(rec);  // при переносе освобождаем старый слот
      const booked = bookSlot(rec);
      if (!booked) return json({ ok: false, error: 'slot_busy_or_missing' });
    } else if (rec.status === 'Отменил запись') {
      clearSlotsOf(rec);
    }
  } finally {
    lock.releaseLock();
  }
  return json({ ok: true });
}

function doGet(e) {
  const p = e.parameter || {};
  if (p.secret !== SECRET) return json({ ok: false, error: 'forbidden' });
  const days = Math.min(Number(p.days) || 7, 31);
  return json({ ok: true, free: freeSlots(days) });
}

// ---------- Запись в слоты ----------

// «10.10.2026 12:30» -> { date: '10.10.2026', time: '12:30' }
function parseVisit(s) {
  const m = String(s).match(/(\d{1,2})\.(\d{1,2})\.(\d{4})\D+(\d{1,2}):(\d{2})/);
  if (!m) return null;
  const pad = (x) => ('0' + x).slice(-2);
  return { date: pad(m[1]) + '.' + pad(m[2]) + '.' + m[3], time: pad(m[4]) + ':' + m[5], month: Number(m[2]) - 1, year: Number(m[3]) };
}

function bookSlot(rec) {
  const v = parseVisit(rec.visit);
  if (!v) return false;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MONTHS[v.month] + ' ' + v.year);
  if (!sheet) return false;
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (fmtDate(rows[i][0]) === v.date && fmtTime(rows[i][2]) === v.time) {
      if (rows[i][3]) return false;  // слот уже занят
      sheet.getRange(i + 1, 4, 1, 4).setValues([[rec.fio, PRIMARY, rec.iin, rec.phone]]);
      return true;
    }
  }
  return false;
}

// Освобождает слоты клиента (ищем по ИИН, если он есть, иначе по ФИО).
function clearSlotsOf(rec) {
  if (!rec.iin && !rec.fio) return;
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach((sheet) => {
    if (sheet.getRange(1, 1).getValue() !== SCHEDULE_HEADERS[0]) return;
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const same = rec.iin ? String(rows[i][5]) === rec.iin : rows[i][3] === rec.fio;
      if (same && rows[i][4] === PRIMARY) sheet.getRange(i + 1, 4, 1, 4).clearContent();
    }
  });
}

function freeSlots(days) {
  const now = new Date();
  const until = new Date(now.getTime() + days * 86400000);
  const result = [];
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach((sheet) => {
    if (sheet.getRange(1, 1).getValue() !== SCHEDULE_HEADERS[0]) return;
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][3]) continue;
      const start = slotStart(rows[i][0], rows[i][2]);
      if (start > now && start < until) {
        result.push(fmtDate(rows[i][0]) + ' ' + rows[i][1] + ' ' + fmtTime(rows[i][2]));
      }
    }
  });
  return result;
}

// ---------- Создание листа на месяц ----------

function createCurrentMonth() {
  const d = new Date();
  createMonth(d.getFullYear(), d.getMonth());
}

function createNextMonth() {
  const d = new Date();
  createMonth(d.getMonth() === 11 ? d.getFullYear() + 1 : d.getFullYear(), (d.getMonth() + 1) % 12);
}

function createMonth(year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = MONTHS[month] + ' ' + year;
  if (ss.getSheetByName(name)) {
    SpreadsheetApp.getUi().alert('Лист «' + name + '» уже есть.');
    return;
  }
  const rows = [];
  for (let d = new Date(year, month, 1); d.getMonth() === month; d.setDate(d.getDate() + 1)) {
    const wd = d.getDay();
    if (wd === 0) continue;  // воскресенье — выходной
    const lastMin = wd === 6 ? 14 * 60 + 30 : 19 * 60 + 30;
    for (let m = 11 * 60; m <= lastMin; m += 30) {
      const t = ('0' + Math.floor(m / 60)).slice(-2) + ':' + ('0' + (m % 60)).slice(-2);
      rows.push([new Date(d), DAYS[wd], t, '', '', '', '', '']);
    }
  }
  const sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, SCHEDULE_HEADERS.length).setValues([SCHEDULE_HEADERS])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#2e7d32');
  sheet.getRange(2, 1, rows.length, SCHEDULE_HEADERS.length).setValues(rows);
  sheet.getRange(2, 1, rows.length, 1).setNumberFormat('dd.MM.yyyy');
  sheet.getRange(2, 3, rows.length, 1).setNumberFormat('@');
  sheet.getRange(2, 6, rows.length, 2).setNumberFormat('@');
  sheet.getRange(2, 8, rows.length, 1).setFormulaR1C1('=IF(R[0]C4="","Свободно","Занято")');
  sheet.getRange(2, 5, rows.length, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList([PRIMARY, 'Повторный приём (1 час)']).build());
  sheet.setFrozenRows(1);
  [12, 14, 8, 34, 30, 15, 16, 13].forEach((w, i) => sheet.setColumnWidth(i + 1, w * 7));
}

// ---------- Вспомогательное ----------

function logRow(rec) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET);
    sheet.appendRow(LOG_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange('C:D').setNumberFormat('@');  // ИИН и телефон как текст
  }
  sheet.appendRow([new Date(), rec.fio, rec.iin, rec.phone, rec.visit, rec.remind, rec.status]);
}

function fmtDate(v) {
  return v instanceof Date ? Utilities.formatDate(v, TZ, 'dd.MM.yyyy') : String(v);
}

function fmtTime(v) {
  return v instanceof Date ? Utilities.formatDate(v, TZ, 'HH:mm') : String(v).slice(0, 5);
}

function slotStart(dateVal, timeVal) {
  const [h, m] = fmtTime(timeVal).split(':').map(Number);
  const d = new Date(dateVal);
  d.setHours(h, m, 0, 0);
  return d;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
