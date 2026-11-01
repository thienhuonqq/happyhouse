// ═══════════════════════════════════════════════════════════════
// Happy House – Google Apps Script Backend
// ═══════════════════════════════════════════════════════════════
// HOW TO DEPLOY:
//   1. Open your Google Sheet → Extensions → Apps Script
//   2. Paste this entire file into Code.gs (replace any existing code)
//   3. Click Save (💾)
//   4. Click Deploy → New deployment (or Manage deployments → edit existing)
//   5. Type: Web app | Execute as: Me | Who has access: Anyone
//   6. Click Deploy → Copy the Web App URL into index.html SHEETS_URL
// ═══════════════════════════════════════════════════════════════

const SHEET_NAME = 'HappyHouseData';
const DELETED_SHEET_NAME = 'HappyHouseDeleted';

const BASE_HEADERS = [
  'id', 'month', 'totalBill', 'totalKwh', 'sharedKwh',
  'pricePerKwh', 'waterBill', 'savedAt'
];
const DEFAULT_PEOPLE = ['Nhi', 'Quỳnh', 'Trang', 'Vy', 'Thuận Ý', 'Chau', 'Khuê', 'Tuyền', 'Trân', 'Như', 'Ngân'];
const HEADERS = BASE_HEADERS.concat(DEFAULT_PEOPLE, ['data_json']);
const MAX_PERSON_COLUMNS = 50;

// ── Get or create the data sheet ──────────────────────────────
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#d9ead3');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(HEADERS.length, 800); // data_json column wide
    sheet.getRange(1, 1, 1, HEADERS.length).setHorizontalAlignment('center');
  }
  // Keep ids as text, otherwise Sheets turns "2026-03" into a date
  sheet.getRange('A:A').setNumberFormat('@');
  ensurePersonColumns(sheet, DEFAULT_PEOPLE);
  return sheet;
}

// ── Log of deleted months, so other devices don't re-upload them ──
function getDeletedSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DELETED_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DELETED_SHEET_NAME);
    sheet.getRange(1, 1, 1, 2).setValues([['id', 'deletedAt']]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  sheet.getRange('A:B').setNumberFormat('@');
  return sheet;
}

function getDeleted() {
  const data = getDeletedSheet().getDataRange().getValues();
  const deleted = {};
  for (let i = 1; i < data.length; i++) deleted[normalizeId(data[i][0])] = String(data[i][1]);
  return deleted;
}

// ── Older rows may hold the id as a Date ──────────────────────
function normalizeId(value) {
  if (value instanceof Date) {
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(value, tz, 'yyyy-MM');
  }
  return String(value);
}

function getHeaders(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  return sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(String);
}

// Add new residents immediately before data_json, preserving every old row.
function ensurePersonColumns(sheet, names) {
  let headers = getHeaders(sheet);
  let jsonIndex = headers.indexOf('data_json');
  if (jsonIndex < 0) {
    jsonIndex = headers.length;
    sheet.getRange(1, jsonIndex + 1).setValue('data_json');
    headers.push('data_json');
  }

  const reserved = new Set(BASE_HEADERS.concat(['data_json', '(Khác)']));
  const available = Math.max(0, MAX_PERSON_COLUMNS - Math.max(0, jsonIndex - BASE_HEADERS.length));
  const missing = [...new Set((names || []).map(String).map(n => n.trim()).filter(n =>
    n && n.length <= 80 && !/^[=+\-@]/.test(n) && !reserved.has(n) && !headers.includes(n)
  ))].slice(0, available);
  if (missing.length) {
    sheet.insertColumnsBefore(jsonIndex + 1, missing.length);
    sheet.getRange(1, jsonIndex + 1, 1, missing.length).setValues([missing]);
    headers.splice(jsonIndex, 0, ...missing);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#d9ead3').setHorizontalAlignment('center');
  }
  sheet.setColumnWidth(headers.indexOf('data_json') + 1, 800);
  return headers;
}

// ── Build a readable row from a record ────────────────────────
function recordToRow(rec, headers) {
  const persons = rec.personResults || [];
  const personTotals = {};
  persons.forEach(p => {
    if (!p.name) return;
    const total = p.total != null ? p.total : (Number(p.dien) || 0) + (Number(p.nuoc) || 0);
    personTotals[p.name] = (personTotals[p.name] || 0) + total;
  });

  return headers.map(header => {
    if (header === 'data_json') return JSON.stringify(rec);
    if (header === 'pricePerKwh') return rec.pricePerKwh ? Math.round(rec.pricePerKwh) : '';
    if (BASE_HEADERS.includes(header)) return rec[header] || '';
    return Object.prototype.hasOwnProperty.call(personTotals, header) ? personTotals[header] : '';
  });
}

// ── GET: return all records as JSON ───────────────────────────
function doGet(e) {
  try {
    const sheet = getSheet();
    const data  = sheet.getDataRange().getValues();
    const records = [];
    const jsonCol = getHeaders(sheet).indexOf('data_json');
    for (let i = 1; i < data.length; i++) {
      try {
        if (data[i][jsonCol]) records.push(JSON.parse(data[i][jsonCol]));
      } catch (err) {}
    }
    return respond({ status: 'ok', records: records, deleted: getDeleted() });
  } catch (err) {
    return respond({ status: 'error', message: err.toString() });
  }
}

// ── POST: upsert or delete ────────────────────────────────────
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const body  = JSON.parse(e.postData.contents);
    const sheet = getSheet();

    if (body.action === 'upsert') {
      const rec  = body.record;
      const headers = ensurePersonColumns(sheet, (rec.personResults || []).map(p => p.name));
      const data = sheet.getDataRange().getValues();
      const row  = recordToRow(rec, headers);
      let found  = false;

      for (let i = 1; i < data.length; i++) {
        if (normalizeId(data[i][0]) === rec.id) {
          sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
          found = true;
          break;
        }
      }

      if (!found) {
        sheet.appendRow(row);
        // Sort by id descending so newest month is at the top
        const last = sheet.getLastRow();
        if (last > 2) {
          sheet.getRange(2, 1, last - 1, headers.length).sort({ column: 1, ascending: false });
        }
      }
      return respond({ status: 'ok' });
    }

    if (body.action === 'delete') {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (normalizeId(data[i][0]) === body.id) {
          sheet.deleteRow(i + 1);
          break;
        }
      }
      getDeletedSheet().appendRow([body.id, new Date().toISOString()]);
      return respond({ status: 'ok' });
    }

    return respond({ status: 'error', message: 'Unknown action: ' + body.action });

  } catch (err) {
    return respond({ status: 'error', message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

// ── JSON response ─────────────────────────────────────────────
function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
