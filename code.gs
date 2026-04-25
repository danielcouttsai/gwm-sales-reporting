// ============================================================
//  GWM Australia – Daily Sales Reporting Tool
//  Google Apps Script Backend  (code.gs)
//
//  SETUP:
//  1. Open Google Sheets → Extensions → Apps Script
//  2. Paste this entire file into Code.gs
//  3. Set SHEET_ID below to your Google Sheet ID
//  4. Deploy as Web App (Execute as: Me, Access: Anyone)
//  5. Copy the Web App URL into GWM_CONFIG.scriptUrl in app.js
// ============================================================

// ── Configuration ────────────────────────────────────────
const SHEET_ID   = 'YOUR_GOOGLE_SHEET_ID_HERE';   // ← Replace with your Sheet ID
const SHEET_NAME = 'submissions';                  // Tab name in your Google Sheet
const ACCESS_CODE = '';                            // Optional: set a shared access code e.g. 'gwm2025'
                                                   // Leave empty to allow all requests

// ── Column order (must match sheet header row exactly) ────
const COLUMNS = [
  'submitted_at',
  'report_date',
  'is_late',
  'dealer_code',
  'dealer_name',
  'region',
  'submitted_by',
  'direction',
  'is_complete_submission',
  'input_method',
  'submission_duration_seconds',
  'last_updated_at',
  'model_bucket',
  'enquiry',
  'test_drives',
  'new_sold',
  'fleet_5_plus',
  'demo_sold',
  'forecast',
];

// ── GET handler – returns rows for dashboard ──────────────
function doGet(e) {
  const params = e.parameter || {};

  // Optional access code check
  if (ACCESS_CODE && params.code !== ACCESS_CODE) {
    return jsonResponse({ error: 'Unauthorized' }, 403);
  }

  try {
    const sheet = getSheet();
    const data  = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return jsonResponse({ rows: [] });
    }

    const headers = data[0].map(h => String(h).trim());
    const rows    = [];

    for (let i = 1; i < data.length; i++) {
      const row = {};
      headers.forEach((h, j) => {
        row[h] = data[i][j];
      });

      // Normalise date before filtering. Google Sheets may return a Date object.
      row['report_date'] = normaliseSheetDate(row['report_date']);

      // Filter by date if provided
      if (params.date && row['report_date'] !== params.date) continue;

      // Filter by dealer if provided
      if (params.dealer && row['dealer_code'] !== params.dealer) continue;

      // Normalise boolean fields
      row['is_late'] = row['is_late'] === true || row['is_late'] === 'TRUE' || row['is_late'] === 1;
      row['is_complete_submission'] = row['is_complete_submission'] === true || row['is_complete_submission'] === 'TRUE' || row['is_complete_submission'] === 1;

      rows.push(row);
    }

    return jsonResponse({ rows });
  } catch (err) {
    return jsonResponse({ error: err.message }, 500);
  }
}

// ── POST handler – appends submission rows ────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const rows = body.rows;

    if (!rows || !Array.isArray(rows) || rows.length === 0) {
      return jsonResponse({ error: 'No rows provided' }, 400);
    }

    // Optional access code check
    if (ACCESS_CODE && body.code !== ACCESS_CODE) {
      return jsonResponse({ error: 'Unauthorized' }, 403);
    }

    const sheet = getSheet();
    ensureHeaders(sheet);

    // Prevent double-sent data. A dealer/date can only have one live submission.
    // If the same dealer submits again for the same report_date, replace the prior model rows.
    const dealerCode = rows[0].dealer_code;
    const reportDate = rows[0].report_date;
    const existingForecasts = getExistingMonthlyForecasts(sheet, dealerCode, reportDate);
    deleteExistingSubmissionRows(sheet, dealerCode, reportDate);

    const appended = [];
    rows.forEach(row => {
      const rowData = COLUMNS.map(col => {
        if (col === 'is_late')    return row[col] ? 'TRUE' : 'FALSE';
        if (col === 'is_complete_submission') return row[col] ? 'TRUE' : 'FALSE';
        if (col === 'fleet_5_plus') return row['fleet'] === '' || row['fleet'] === null || row['fleet'] === undefined ? '' : safeInt(row['fleet']);   // map fleet → fleet_5_plus
        if (col === 'forecast') {
          if (row[col] === '' || row[col] === null || row[col] === undefined) {
            const model = String(row.model_bucket || '').trim();
            return existingForecasts[model] !== undefined ? existingForecasts[model] : '';
          }
          return safeInt(row[col]);
        }
        return row[col] !== undefined ? row[col] : '';
      });
      sheet.appendRow(rowData);
      appended.push(row.model_bucket);
    });

    return jsonResponse({
      success: true,
      appended: appended.length,
      message: `${appended.length} rows written for dealer ${rows[0].dealer_code}`,
    });

  } catch (err) {
    return jsonResponse({ error: err.message }, 500);
  }
}

// ── Sheet helpers ─────────────────────────────────────────
function getSheet() {
  const ss = SHEET_ID
    ? SpreadsheetApp.openById(SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  return sheet;
}

function ensureHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLUMNS);
  } else {
    const existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), COLUMNS.length)).getValues()[0].map(h => String(h || '').trim());
    const needsUpgrade = COLUMNS.some((col, i) => existing[i] !== col);
    if (needsUpgrade) {
      sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
    }
  }

  // Style the header row and keep widths consistent.
  const headerRange = sheet.getRange(1, 1, 1, COLUMNS.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#111111');
  headerRange.setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 180);  // submitted_at
  sheet.setColumnWidth(2, 110);  // report_date
  sheet.setColumnWidth(5, 200);  // dealer_name
  sheet.setColumnWidth(9, 160);  // is_complete_submission
  sheet.setColumnWidth(10, 130); // input_method
  sheet.setColumnWidth(12, 180); // last_updated_at
  sheet.setColumnWidth(13, 160); // model_bucket
}


function getExistingMonthlyForecasts(sheet, dealerCode, reportDate) {
  const forecasts = {};
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return forecasts;

  const data = sheet.getRange(2, 1, lastRow - 1, COLUMNS.length).getValues();
  const dealerCol = COLUMNS.indexOf('dealer_code');
  const dateCol = COLUMNS.indexOf('report_date');
  const modelCol = COLUMNS.indexOf('model_bucket');
  const forecastCol = COLUMNS.indexOf('forecast');
  const submittedCol = COLUMNS.indexOf('submitted_at');
  const targetMonth = String(reportDate || '').slice(0, 7);
  const latest = {};

  for (let i = 0; i < data.length; i++) {
    const rowDealer = String(data[i][dealerCol] || '').trim();
    const rowDate = normaliseSheetDate(data[i][dateCol]);
    if (rowDealer !== String(dealerCode).trim()) continue;
    if (String(rowDate).slice(0, 7) !== targetMonth) continue;

    const model = String(data[i][modelCol] || '').trim();
    const forecast = data[i][forecastCol];
    if (!model || forecast === '' || forecast === null || forecast === undefined) continue;

    const ts = new Date(data[i][submittedCol] || 0).getTime() || 0;
    if (!latest[model] || ts >= latest[model].ts) {
      latest[model] = { ts: ts, val: safeInt(forecast) };
    }
  }

  Object.keys(latest).forEach(model => forecasts[model] = latest[model].val);
  return forecasts;
}

function deleteExistingSubmissionRows(sheet, dealerCode, reportDate) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const data = sheet.getRange(2, 1, lastRow - 1, COLUMNS.length).getValues();
  const dealerCol = COLUMNS.indexOf('dealer_code');
  const dateCol   = COLUMNS.indexOf('report_date');

  // Delete from bottom up so row numbers stay valid.
  for (let i = data.length - 1; i >= 0; i--) {
    const rowDealer = String(data[i][dealerCol] || '').trim();
    const rowDate   = normaliseSheetDate(data[i][dateCol]);
    if (rowDealer === String(dealerCode).trim() && rowDate === String(reportDate).trim()) {
      sheet.deleteRow(i + 2);
    }
  }
}

function normaliseSheetDate(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value || '').trim();
}

function safeInt(val) {
  const n = parseInt(val, 10);
  return isNaN(n) || n < 0 ? 0 : n;
}

// ── JSON response helper ──────────────────────────────────
function jsonResponse(data, status) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ── Test function (run manually in Apps Script editor) ────
function testSetup() {
  const sheet = getSheet();
  ensureHeaders(sheet);
  Logger.log('Sheet ready: ' + sheet.getName() + ' — rows: ' + sheet.getLastRow());
}
