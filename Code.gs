// ═══════════════════════════════════════════════════════════════
// Happy House – Google Apps Script Backend
// ═══════════════════════════════════════════════════════════════
// HOW TO DEPLOY:
//   1. Open your Google Sheet → Extensions → Apps Script
//   2. Paste this entire file into Code.gs (replace any existing code)
//   3. Click Save (💾)
//   4. Click Deploy → New deployment
//   5. Type: Web app
//   6. Execute as: Me
//   7. Who has access: Anyone
//   8. Click Deploy → Copy the Web App URL
//   9. Paste that URL into index.html where it says SHEETS_URL = ''
// ═══════════════════════════════════════════════════════════════

const SHEET_NAME = 'HappyHouseData';

// ── Get or create the data sheet ──────────────────────────────
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, 3).setValues([['id', 'savedAt', 'data_json']]);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 100);
    sheet.setColumnWidth(2, 180);
    sheet.setColumnWidth(3, 600);
  }
  return sheet;
}

// ── GET: load all records ─────────────────────────────────────
function doGet(e) {
  try {
    const sheet = getSheet();
    const data  = sheet.getDataRange().getValues();
    const records = [];
    for (let i = 1; i < data.length; i++) {
      try {
        if (data[i][2]) records.push(JSON.parse(data[i][2]));
      } catch (err) {
        // skip malformed row
      }
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
      let found  = false;

      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === rec.id) {
          sheet.getRange(i + 1, 1, 1, 3).setValues([
            [rec.id, rec.savedAt || new Date().toISOString(), JSON.stringify(rec)]
          ]);
          found = true;
          break;
        }
      }

      if (!found) {
        sheet.appendRow([
          rec.id,
          rec.savedAt || new Date().toISOString(),
          JSON.stringify(rec)
        ]);
        // Sort rows by id descending (newest month first), keeping header
        sortSheet(sheet);
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

// ── Sort sheet by id column descending ───────────────────────
function sortSheet(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 2) return;
  const range = sheet.getRange(2, 1, lastRow - 1, 3);
  range.sort({ column: 1, ascending: false });
}

// ── Helper: JSON response with CORS headers ───────────────────
function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
