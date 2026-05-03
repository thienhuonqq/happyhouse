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

const HEADERS = [
  'id', 'month', 'totalBill', 'totalKwh', 'sharedKwh',
  'pricePerKwh', 'waterBill', 'savedAt',
  'Nhi', 'Quỳnh', 'Trang', 'Vy', 'Thuận Ý', 'Chau', 'Khuê', 'Tuyền', 'Trân',
  'data_json'
];

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
  return sheet;
}

// ── Build a readable row from a record ────────────────────────
function recordToRow(rec) {
  const persons = rec.personResults || [];
  function getPersonTotal(name) {
    const p = persons.find(p => p.name === name);
    return p ? (p.total || 0) : '';
  }
  return [
    rec.id || '',
    rec.month || '',
    rec.totalBill || '',
    rec.totalKwh || '',
    rec.sharedKwh || '',
    rec.pricePerKwh ? Math.round(rec.pricePerKwh) : '',
    rec.waterBill || '',
    rec.savedAt || '',
    getPersonTotal('Nhi'),
    getPersonTotal('Quỳnh'),
    getPersonTotal('Trang'),
    getPersonTotal('Vy'),
    getPersonTotal('Thuận Ý'),
    getPersonTotal('Chau'),
    getPersonTotal('Khuê'),
    getPersonTotal('Tuyền'),
    getPersonTotal('Trân'),
    JSON.stringify(rec)
  ];
}

// ── GET: return all records as JSON ───────────────────────────
function doGet(e) {
  try {
    const sheet = getSheet();
    const data  = sheet.getDataRange().getValues();
    const records = [];
    const jsonCol = HEADERS.length - 1; // last column = data_json
    for (let i = 1; i < data.length; i++) {
      try {
        if (data[i][jsonCol]) records.push(JSON.parse(data[i][jsonCol]));
      } catch (err) {}
    }
    return respond({ status: 'ok', records: records });
  } catch (err) {
    return respond({ status: 'error', message: err.toString() });
  }
}

// ── POST: upsert or delete ────────────────────────────────────
function doPost(e) {
  try {
    const body  = JSON.parse(e.postData.contents);
    const sheet = getSheet();

    if (body.action === 'upsert') {
      const rec  = body.record;
      const data = sheet.getDataRange().getValues();
      const row  = recordToRow(rec);
      let found  = false;

      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === rec.id) {
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
          sheet.getRange(2, 1, last - 1, HEADERS.length).sort({ column: 1, ascending: false });
        }
      }
      return respond({ status: 'ok' });
    }

    if (body.action === 'delete') {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === body.id) {
          sheet.deleteRow(i + 1);
          break;
        }
      }
      return respond({ status: 'ok' });
    }

    return respond({ status: 'error', message: 'Unknown action: ' + body.action });

  } catch (err) {
    return respond({ status: 'error', message: err.toString() });
  }
}

// ── JSON response ─────────────────────────────────────────────
function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
