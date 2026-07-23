/**
 * Backend for the Library Book Inventory page (library.html).
 * Deploy this as a Web App bound to a Google Sheet - see setup steps
 * in the project chat / commit message for the copy-paste instructions.
 */

const SHEET_NAME = 'Books';
const HEADERS = ['ID', 'Name', 'Genre', 'Shelf', 'Status', 'Borrower', 'DueDate', 'DateAdded', 'Notes'];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

function doGet(e) {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const books = data.slice(1)
    .filter(function (row) { return row[0]; })
    .map(function (row) {
      const obj = {};
      headers.forEach(function (h, i) { obj[keyFor_(h)] = row[i]; });
      return obj;
    });
  return respond_({ ok: true, books: books });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    const sheet = getSheet_();

    if (body.action === 'add') {
      const id = nextId_(sheet);
      const b = body.book || {};
      sheet.appendRow([
        id,
        b.name || '',
        b.genre || '',
        b.shelf || '',
        b.status || 'Available',
        b.borrower || '',
        b.dueDate || '',
        todayStr_(),
        b.notes || ''
      ]);
      return respond_({ ok: true, id: id });
    }

    if (body.action === 'update') {
      const b = body.book || {};
      const rowIndex = findRowById_(sheet, b.id);
      if (rowIndex === -1) return respond_({ ok: false, error: 'Book not found' });
      const existing = sheet.getRange(rowIndex, 1, 1, HEADERS.length).getValues()[0];
      const updated = HEADERS.map(function (h, i) {
        const key = keyFor_(h);
        return (b[key] !== undefined && b[key] !== null) ? b[key] : existing[i];
      });
      sheet.getRange(rowIndex, 1, 1, HEADERS.length).setValues([updated]);
      return respond_({ ok: true });
    }

    if (body.action === 'delete') {
      const rowIndex = findRowById_(sheet, body.id);
      if (rowIndex === -1) return respond_({ ok: false, error: 'Book not found' });
      sheet.deleteRow(rowIndex);
      return respond_({ ok: true });
    }

    return respond_({ ok: false, error: 'Unknown action: ' + body.action });
  } catch (err) {
    return respond_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function keyFor_(header) {
  return header.charAt(0).toLowerCase() + header.slice(1);
}

function todayStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function nextId_(sheet) {
  const data = sheet.getDataRange().getValues();
  let max = 0;
  for (let i = 1; i < data.length; i++) {
    const m = String(data[i][0] || '').match(/^vriv-(\d+)$/);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  }
  return 'vriv-' + String(max + 1).padStart(5, '0');
}

function findRowById_(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
