/**
 * Приёмник результатов квеста «Тайна Королевства Вердан»
 * Вставьте этот код в Apps Script вашей Google-таблицы
 * (Расширения → Apps Script), затем разверните как веб-приложение.
 * Подробная инструкция: docs/googlesheets-setup.md
 */

var TOKEN = 'verdania-2026'; // должен совпадать со значением TOKEN в app.js квеста

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.token !== TOKEN) {
      return json({ ok: false, error: 'Неверный ключ' });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Результаты') || ss.getSheets()[0];
    ensureHeader(sheet);

    var ts = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd.MM.yyyy HH:mm:ss');
    sheet.appendRow([
      ts,
      String(data.name || '').slice(0, 100),
      String(data.group || '').slice(0, 30),
      data.score,
      data.maxScore,
      data.time,
      data.seconds
    ]);

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function doGet() {
  return json({ ok: true, info: 'Приёмник результатов квеста работает' });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ensureHeader(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Дата и время', 'ФИО', 'Группа', 'Балл', 'Максимум', 'Время (мм:сс)', 'Время (сек)']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#e7d3a9');
    sheet.setFrozenRows(1);
  }
}
