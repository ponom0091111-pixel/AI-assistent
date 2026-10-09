// Связка бота Карины с Google Таблицей.
//
// Лист расписания на месяц (например «Октябрь 2026»): каждая строка — слот 30 минут.
// Столбцы:
//   A Дата | B День недели | C Начало | D Окончание | E ФИО | F Процедура | G ИИН |
//   H Телефон | I Источник | J Статус записи | K Создано | L Статус слота
//
// Слот занят, если в нём есть ФИО или статус записи (Бронь / Записан),
// а также если в предыдущем слоте того же дня стоит повторный приём на 1 час.
// «Бронь» — время держится за клиентом до оплаты, HOLD_MINUTES минут;
// после этого неоплаченная бронь считается истёкшей и слот снова свободен.
//
// POST (doPost) — бот присылает JSON из блока <crm>...</crm>:
//   • «Бронь»          — занять слот до оплаты (как только клиент выбрал время);
//   • «Записан»        — подтвердить запись в слоте (после оплаты);
//   • «Перенёс запись» — освободить старый слот клиента и занять новый;
//   У клиента всегда одна запись: при любой новой брони/записи старая удаляется.
//   • «Отменил запись» — освободить слот;
//   • любой статус дополнительно пишется строкой в лист «Клиенты» (журнал).
//   Ответ: {"ok":true,"start":"12:30","end":"13:00"} или {"ok":false,"error":"slot_busy"}.
//   Перенос оплаченной записи: reschedule_limit (уже переносили) или too_late (меньше суток) — бронь сгорает.
// GET (doGet) — свободные слоты для бота: ?secret=...&days=7
//   Ответ: {"recommended": [окна рядом с уже записанными], "free": [все свободные окна]}
//
// createCurrentMonth / createNextMonth — создать лист расписания на месяц
// (меню «Расписание» в таблице).

const SECRET = 'ЗАМЕНИТЕ_НА_СВОЙ_СЕКРЕТНЫЙ_КЛЮЧ';
const HOLD_MINUTES = 60;
const SLOT_MINUTES = 30;
const LOG_SHEET = 'Клиенты';
const LOG_HEADERS = ['Дата и время', 'ФИО', 'ИИН', 'Телефон', 'Источник', 'Начало визита',
  'Окончание визита', 'Напомнить', 'Категория', 'Статус', 'Переносов'];
// Правила оплаченной брони: перенос — только один раз и не позднее чем за сутки до визита.
const MAX_RESCHEDULES = 1;
const RESCHEDULE_MIN_HOURS = 24;
const SCHEDULE_HEADERS = ['Дата', 'День недели', 'Начало', 'Окончание', 'ФИО', 'Процедура', 'ИИН',
  'Телефон', 'Источник', 'Статус записи', 'Создано', 'Статус слота'];
const SOURCES = ['Instagram — платная реклама', 'Instagram — бесплатно', 'Facebook — платная реклама',
  'Facebook — бесплатно', '2ГИС', 'Рекомендация', 'Другое', 'Не указан'];
const PRIMARY = 'Первичная диагностика ЖКТ';
const REPEAT = 'Повторный приём (1 час)';
const DAYS = ['Воскресенье', 'Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота'];
const MONTHS = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль',
  'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
const TZ = 'Asia/Almaty';

// Индексы столбцов (с нуля) в листе расписания.
const C = { date: 0, day: 1, start: 2, end: 3, fio: 4, proc: 5, iin: 6, phone: 7, source: 8, status: 9, created: 10 };
// Сколько столбцов занимает запись клиента: от ФИО до «Создано».
const REC_COLS = C.created - C.fio + 1;

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Расписание')
    .addItem('Создать лист на текущий месяц', 'createCurrentMonth')
    .addItem('Создать лист на следующий месяц', 'createNextMonth')
    .addItem('Очистить истёкшие брони', 'clearExpiredHolds')
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
    phone: normPhone(data['Телефон']),
    visit: data['Дата визита'] || '',
    remind: data['Напомнить'] || '',
    category: data['Категория'] || '',
    source: data['Источник'] || '',
    status: data['Статус'] || '',
  };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let result = { ok: true };
  try {
    if (rec.status === 'Перенёс запись') {
      result = checkReschedule(rec);
      if (result.ok) result = moveClientTo(rec, 'Записан');
    } else if (rec.status === 'Бронь' || rec.status === 'Записан') {
      result = moveClientTo(rec, rec.status);
    } else if (rec.status === 'Отменил запись') {
      clearSlotsOf(rec);
    }
    // Журнал обновляем только при успехе, чтобы отказ (занято, перенос запрещён) не менял данные клиента.
    if (result.ok) {
      const v = parseVisit(rec.visit);
      logRow(rec, v ? v.time : '', v ? addMinutes(v.time, SLOT_MINUTES) : '');
    }
  } finally {
    lock.releaseLock();
  }
  return json(result);
}

function doGet(e) {
  const p = e.parameter || {};
  if (p.secret !== SECRET) return json({ ok: false, error: 'forbidden' });
  const days = Math.min(Number(p.days) || 7, 31);
  const slots = freeSlots(days);
  return json({
    ok: true,
    // Сначала предлагаем окна рядом с уже записанными клиентами, чтобы записи шли плотно.
    recommended: slots.filter((s) => s.near).map((s) => s.text),
    free: slots.map((s) => s.text),
  });
}

// ---------- Слоты ----------

// «10.10.2026 12:30» -> { date: '10.10.2026', time: '12:30', month: 9, year: 2026 }
function parseVisit(s) {
  const m = String(s).match(/(\d{1,2})\.(\d{1,2})\.(\d{4})\D+(\d{1,2}):(\d{2})/);
  if (!m) return null;
  return { date: pad(m[1]) + '.' + pad(m[2]) + '.' + m[3], time: pad(m[4]) + ':' + m[5],
    month: Number(m[2]) - 1, year: Number(m[3]) };
}

function findSlot(v) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MONTHS[v.month] + ' ' + v.year);
  if (!sheet) return null;
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (fmtDate(rows[i][C.date]) === v.date && fmtTime(rows[i][C.start]) === v.time) {
      return { sheet: sheet, rows: rows, i: i };
    }
  }
  return null;
}

// Слот занят кем-то другим? Свою бронь клиент может подтвердить.
function isBusy(rows, i, rec) {
  const r = rows[i];
  const prev = rows[i - 1];
  if (prev && i > 1 && prev[C.proc] === REPEAT && fmtDate(prev[C.date]) === fmtDate(r[C.date])) {
    return true;  // вторая половина часового повторного приёма
  }
  if (!r[C.fio] && !r[C.status]) return false;
  if (rec && sameClient(r, rec)) return false;  // своё окно клиент может обновить (например, дослать ФИО и ИИН)
  if (r[C.status] === 'Бронь' && holdExpired(r)) return false;
  return true;
}

function bookSlot(rec, status) {
  const v = parseVisit(rec.visit);
  if (!v) return { ok: false, error: 'bad_visit_date' };
  const slot = findSlot(v);
  if (!slot) return { ok: false, error: 'slot_missing' };
  if (isBusy(slot.rows, slot.i, rec)) return { ok: false, error: 'slot_busy' };

  const r = slot.rows[slot.i];
  // При подтверждении сохраняем данные из брони, если в новом сообщении их нет.
  const fio = rec.fio || (sameClient(r, rec) ? r[C.fio] : '');
  const phone = rec.phone || (sameClient(r, rec) ? String(r[C.phone]) : '');
  const source = rec.source || (sameClient(r, rec) ? r[C.source] : '');
  slot.sheet.getRange(slot.i + 1, C.fio + 1, 1, REC_COLS)
    .setValues([[fio, PRIMARY, rec.iin, phone, source, status, new Date()]]);
  return { ok: true, start: v.time, end: addMinutes(v.time, SLOT_MINUTES) };
}

// Освобождает слоты клиента (ищем по телефону, ИИН или ФИО).
function clearSlotsOf(rec) {
  if (!rec.phone && !rec.iin && !rec.fio) return;
  scheduleSheets().forEach((sheet) => {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][C.proc] === PRIMARY && sameClient(rows[i], rec)) {
        sheet.getRange(i + 1, C.fio + 1, 1, REC_COLS).clearContent();
      }
    }
  });
}

// Перенос оплаченной записи: не больше MAX_RESCHEDULES раз и не позднее чем за сутки до визита.
// Ответ с ошибкой reschedule_limit или too_late означает, что бронь сгорает.
function checkReschedule(rec) {
  let visitStart = null;
  scheduleSheets().forEach((sheet) => {
    sheet.getDataRange().getValues().forEach((row, i) => {
      if (i > 0 && row[C.status] === 'Записан' && sameClient(row, rec)) {
        visitStart = slotStart(row[C.date], row[C.start]);
      }
    });
  });
  if (!visitStart) return { ok: true };  // оплаченной записи нет — переносить нечего
  const log = findLogRow(rec);
  const done = log ? Number(log.row[10]) || 0 : 0;
  if (done >= MAX_RESCHEDULES) return { ok: false, error: 'reschedule_limit' };
  if (visitStart.getTime() - Date.now() < RESCHEDULE_MIN_HOURS * 3600000) {
    return { ok: false, error: 'too_late' };
  }
  return { ok: true };
}

// У клиента в расписании всегда одна запись: новая заменяет старую.
// Сначала проверяем, что новое время свободно, и только потом удаляем старую запись
// клиента (бронь или запись на другое время/день) и занимаем новый слот.
function moveClientTo(rec, status) {
  const v = parseVisit(rec.visit);
  if (!v) return { ok: false, error: 'bad_visit_date' };
  const target = findSlot(v);
  if (!target) return { ok: false, error: 'slot_missing' };
  if (isBusy(target.rows, target.i, rec)) return { ok: false, error: 'slot_busy' };
  fillFromExisting(rec);  // переносим ФИО, ИИН, телефон и источник со старой записи
  clearSlotsOf(rec);
  return bookSlot(rec, status);
}

// Дополняет пустые поля клиента данными из его текущей записи в расписании.
function fillFromExisting(rec) {
  scheduleSheets().forEach((sheet) => {
    sheet.getDataRange().getValues().forEach((row, i) => {
      if (i === 0 || row[C.proc] !== PRIMARY || !sameClient(row, rec)) return;
      rec.fio = rec.fio || row[C.fio];
      rec.iin = rec.iin || String(row[C.iin] || '');
      rec.phone = rec.phone || normPhone(row[C.phone]);
      rec.source = rec.source || row[C.source];
    });
  });
}

function clearExpiredHolds() {
  scheduleSheets().forEach((sheet) => {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][C.status] === 'Бронь' && holdExpired(rows[i])) {
        sheet.getRange(i + 1, C.fio + 1, 1, REC_COLS).clearContent();
      }
    }
  });
}

// Свободные слоты на ближайшие days дней.
// near = соседний слот того же дня уже занят (запись, бронь или повторный приём).
// Подпись (день) — начало до 15:00, (вечер) — с 15:00.
function freeSlots(days) {
  const now = new Date();
  const until = new Date(now.getTime() + days * 86400000);
  const result = [];
  scheduleSheets().forEach((sheet) => {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (isBusy(rows, i, null)) continue;
      const start = slotStart(rows[i][C.date], rows[i][C.start]);
      if (!(start > now && start < until)) continue;
      const day = fmtDate(rows[i][C.date]);
      const sameDayBusy = (j) => j >= 1 && j < rows.length && fmtDate(rows[j][C.date]) === day && isBusy(rows, j, null);
      const t = fmtTime(rows[i][C.start]);
      result.push({
        text: day + ' ' + rows[i][C.day] + ' ' + t + '–' + addMinutes(t, SLOT_MINUTES) +
          (Number(t.slice(0, 2)) < 15 ? ' (день)' : ' (вечер)'),
        near: sameDayBusy(i - 1) || sameDayBusy(i + 1),
      });
    }
  });
  return result;
}

function sameClient(row, rec) {
  if (rec.phone && normPhone(row[C.phone]) === rec.phone) return true;
  if (rec.iin && String(row[C.iin]) === rec.iin) return true;
  if (!rec.phone && !rec.iin && rec.fio && row[C.fio] === rec.fio) return true;
  return false;
}

function holdExpired(row) {
  const created = row[C.created];
  if (!(created instanceof Date)) return false;
  return Date.now() - created.getTime() > HOLD_MINUTES * 60000;
}

function scheduleSheets() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .filter((s) => s.getRange(1, 1).getValue() === SCHEDULE_HEADERS[0]);
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
    for (let m = 11 * 60; m <= lastMin; m += SLOT_MINUTES) {
      const t = pad(Math.floor(m / 60)) + ':' + pad(m % 60);
      rows.push([new Date(d), DAYS[wd], t, addMinutes(t, SLOT_MINUTES), '', '', '', '', '', '', '', '']);
    }
  }
  const n = rows.length;
  const sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, SCHEDULE_HEADERS.length).setValues([SCHEDULE_HEADERS])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#2e7d32');
  sheet.getRange(2, 3, n, 2).setNumberFormat('@');
  sheet.getRange(2, 7, n, 2).setNumberFormat('@');
  sheet.getRange(2, 1, n, SCHEDULE_HEADERS.length).setValues(rows);
  sheet.getRange(2, 1, n, 1).setNumberFormat('dd.MM.yyyy');
  sheet.getRange(2, 11, n, 1).setNumberFormat('dd.MM.yyyy HH:mm');
  sheet.getRange(2, 12, n, 1).setFormulaR1C1(
    '=IF(OR(R[0]C5<>"",R[0]C10<>""),"Занято",IF(AND(R[-1]C6="' + REPEAT + '",R[-1]C1=R[0]C1),"Занято","Свободно"))');
  sheet.getRange(2, 6, n, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList([PRIMARY, REPEAT]).build());
  sheet.getRange(2, 9, n, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(SOURCES).build());
  sheet.getRange(2, 10, n, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Бронь', 'Записан']).build());
  sheet.setFrozenRows(1);
  [12, 14, 8, 10, 32, 28, 15, 16, 22, 14, 16, 13].forEach((w, i) => sheet.setColumnWidth(i + 1, w * 7));
}

// ---------- Вспомогательное ----------

// Журнал «Клиенты»: по одной строке на клиента (ищем по телефону, ИИН или ФИО).
// Новые данные обновляют строку клиента, а не добавляют новую — хранится только актуальное.
function logSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET);
    sheet.appendRow(LOG_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange('C:D').setNumberFormat('@');  // ИИН и телефон как текст
  }
  return sheet;
}

function findLogRow(rec) {
  const sheet = logSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if ((rec.phone && normPhone(r[3]) === rec.phone) || (rec.iin && String(r[2]) === rec.iin) ||
        (!rec.phone && !rec.iin && rec.fio && r[1] === rec.fio)) {
      return { sheet: sheet, idx: i, row: r };
    }
  }
  return null;
}

function logRow(rec, start, end) {
  const found = findLogRow(rec);
  const sheet = found ? found.sheet : logSheet();
  const v = parseVisit(rec.visit);
  const old = found ? found.row : ['', '', '', '', '', '', '', '', '', '', 0];
  const keep = (val, i) => val || old[i] || '';
  const cancelled = rec.status === 'Отменил запись';
  // Счётчик переносов: новая бронь начинает его заново, успешный перенос увеличивает.
  let moves = Number(old[10]) || 0;
  if (rec.status === 'Бронь') moves = 0;
  if (rec.status === 'Перенёс запись') moves += 1;
  const row = [new Date(), keep(rec.fio, 1), keep(rec.iin, 2), keep(rec.phone, 3), keep(rec.source, 4),
    cancelled ? '' : (v ? v.date + ' ' + start : old[5] || ''),
    cancelled ? '' : (v ? v.date + ' ' + end : old[6] || ''),
    rec.remind || old[7] || '', keep(rec.category, 8), rec.status || old[9] || '', moves];
  if (found) sheet.getRange(found.idx + 1, 1, 1, row.length).setValues([row]);
  else sheet.appendRow(row);
}

function normPhone(p) {
  const d = String(p || '').replace(/\D/g, '');
  if (!d) return '';
  return d.length === 11 && d[0] === '8' ? '7' + d.slice(1) : d;
}

function addMinutes(hhmm, minutes) {
  const [h, m] = String(hhmm).split(':').map(Number);
  const total = h * 60 + m + minutes;
  return pad(Math.floor(total / 60)) + ':' + pad(total % 60);
}

function pad(x) {
  return ('0' + x).slice(-2);
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
