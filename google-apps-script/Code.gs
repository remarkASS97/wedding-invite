const SHEET_NAME = 'Ответы';

function doPost(e) {
  const sheet = getSheet_();
  const data = e.parameter || {};

  sheet.appendRow([
    new Date(),
    data.name || '',
    data.attendance || '',
    data.drinks || '',
    data.food || '',
    data.message || '',
    data.submittedAt || '',
    data.source || '',
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Дата записи',
      'Имя',
      'Присутствие',
      'Напитки',
      'Пищевые ограничения',
      'Комментарий',
      'Время отправки с сайта',
      'Страница',
    ]);
  }

  return sheet;
}
