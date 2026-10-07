// Приём данных от бота Карины и запись строки в Google Таблицу.
// Бот (платформа) отправляет POST-запрос с JSON из блока <crm>...</crm>.

const SECRET = 'ЗАМЕНИТЕ_НА_СВОЙ_СЕКРЕТНЫЙ_КЛЮЧ';
const SHEET_NAME = 'Клиенты';
const HEADERS = ['Дата и время', 'ФИО', 'ИИН', 'Телефон', 'Дата визита', 'Напомнить', 'Статус'];

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  if (data.secret !== SECRET) {
    return json({ ok: false, error: 'forbidden' });
  }

  const sheet = getSheet();
  const iin = String(data['ИИН'] || '').replace(/\D/g, '');
  sheet.appendRow([
    new Date(),
    data['ФИО'] || '',
    iin.length === 12 ? iin : '',
    data['Телефон'] || '',
    data['Дата визита'] || '',
    data['Напомнить'] || '',
    data['Статус'] || '',
  ]);
  return json({ ok: true });
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    // ИИН и телефон храним как текст, чтобы не потерять ведущие нули.
    sheet.getRange('C:D').setNumberFormat('@');
  }
  return sheet;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
