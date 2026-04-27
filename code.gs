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
const SHEET_ID   = '1RgHHcIgq_onYKgh9PC6IqsFXxTWbgM_GAVxJSV0NPrY';   // Connected Google Sheet ID
const SHEET_NAME = 'submissions';                  // Tab name in your Google Sheet
const CONTROL_SHEET_NAME = 'submission_controls';   // OEM reopen / lock control tab
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


const DAILY_VALUE_COLUMNS = ['enquiry','test_drives','new_sold','fleet_5_plus','demo_sold'];
const MIN_EXPECTED_MODEL_ROWS = 19;
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

      const rowReportDate = normaliseSheetDate(row['report_date']);
      row['report_date'] = rowReportDate;

      // Filter by date if provided
      if (params.date && rowReportDate !== String(params.date).trim()) continue;

      // Filter by dealer if provided
      if (params.dealer && String(row['dealer_code']).trim() !== String(params.dealer).trim()) continue;

      // Normalise boolean fields
      row['is_late'] = row['is_late'] === true || row['is_late'] === 'TRUE' || row['is_late'] === 1;
      row['is_complete_submission'] = row['is_complete_submission'] === true || row['is_complete_submission'] === 'TRUE' || row['is_complete_submission'] === 1;

      rows.push(row);
    }

    const payload = { rows };
    if (params.include_controls === '1' && params.date) {
      payload.unlockedDealerCodes = getUnlockedDealerCodes(params.date);
    }
    return jsonResponse(payload);
  } catch (err) {
    return jsonResponse({ error: err.message }, 500);
  }
}

// ── POST handler – appends submission rows ────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    // OEM dashboard control actions.
    if (body.action === 'unlock_submission') {
      if (ACCESS_CODE && body.code !== ACCESS_CODE) {
        return jsonResponse({ error: 'Unauthorized' }, 403);
      }
      recordSubmissionUnlock(body.dealer_code, body.report_date, body.unlocked_by || 'OEM Dashboard');
      return jsonResponse({ success: true, action: 'unlock_submission', dealer_code: body.dealer_code, report_date: body.report_date });
    }

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

    const dealerCode = String(rows[0].dealer_code || '').trim();
    const reportDate = normaliseSheetDate(rows[0].report_date);
    validateSubmissionRows(rows, dealerCode, reportDate);

    const alreadySubmitted = hasExistingSubmissionRows(sheet, dealerCode, reportDate);
    const isUnlocked = isSubmissionUnlocked(dealerCode, reportDate);
    if (alreadySubmitted && !isUnlocked) {
      return jsonResponse({
        success: false,
        error: 'Submission already exists for this dealer/date. Reopen it from the dashboard before resubmitting.',
        code: 'SUBMISSION_LOCKED',
        dealer_code: dealerCode,
        report_date: reportDate,
      }, 409);
    }

    // If reopened, resubmission replaces the old dealer/date rows and then locks the dealer again.
    const existingForecasts = getExistingMonthlyForecasts(sheet, dealerCode, reportDate);
    if (alreadySubmitted) deleteExistingSubmissionRows(sheet, dealerCode, reportDate);

    const appended = [];
    rows.forEach(row => {
      const isComplete = isServerCompleteSubmission(rows, dealerCode, reportDate);
      const rowData = COLUMNS.map(col => {
        if (col === 'is_late') return row[col] ? 'TRUE' : 'FALSE';
        if (col === 'is_complete_submission') return isComplete ? 'TRUE' : 'FALSE';
        if (col === 'report_date') return reportDate;
        if (col === 'dealer_code') return dealerCode;
        if (col === 'fleet_5_plus') return coerceDailyValue(row['fleet_5_plus'] !== undefined ? row['fleet_5_plus'] : row['fleet']);
        if (DAILY_VALUE_COLUMNS.indexOf(col) >= 0) return coerceDailyValue(row[col]);
        if (col === 'forecast') {
          if (row[col] === '' || row[col] === null || row[col] === undefined) {
            const model = String(row.model_bucket || '').trim();
            return existingForecasts[model] !== undefined ? existingForecasts[model] : '';
          }
          return safeInt(row[col]);
        }
        if (col === 'last_updated_at') return row[col] || row['submitted_at'] || new Date().toISOString();
        return row[col] !== undefined ? row[col] : '';
      });
      sheet.appendRow(rowData);
      appended.push(row.model_bucket);
    });

    clearSubmissionUnlock(dealerCode, reportDate);

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
  const lastRow = sheet.getLastRow();

  if (lastRow === 0) {
    sheet.appendRow(COLUMNS);
    styleHeader(sheet);
    return;
  }

  const currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), COLUMNS.length)).getValues()[0]
    .map(h => String(h || '').trim())
    .filter(h => h !== '');
  const matches = COLUMNS.length === currentHeaders.length && COLUMNS.every((h, i) => h === currentHeaders[i]);
  if (matches) {
    styleHeader(sheet);
    return;
  }

  // Safe migration for older sheets. Rebuild rows by header name so existing historical data is preserved.
  const existingValues = sheet.getDataRange().getValues();
  const oldHeaders = existingValues[0].map(h => String(h || '').trim());
  const migrated = [COLUMNS];

  for (let r = 1; r < existingValues.length; r++) {
    const obj = {};
    oldHeaders.forEach((h, c) => { if (h) obj[h] = existingValues[r][c]; });

    migrated.push(COLUMNS.map(col => {
      if (col === 'fleet_5_plus') return obj['fleet_5_plus'] !== undefined ? obj['fleet_5_plus'] : (obj['fleet'] !== undefined ? obj['fleet'] : '');
      if (col === 'is_complete_submission') return obj[col] !== undefined ? obj[col] : '';
      if (col === 'input_method') return obj[col] !== undefined ? obj[col] : '';
      if (col === 'submission_duration_seconds') return obj[col] !== undefined ? obj[col] : '';
      if (col === 'last_updated_at') return obj[col] !== undefined ? obj[col] : (obj['submitted_at'] || '');
      return obj[col] !== undefined ? obj[col] : '';
    }));
  }

  sheet.clearContents();
  sheet.getRange(1, 1, migrated.length, COLUMNS.length).setValues(migrated);
  styleHeader(sheet);
}

function styleHeader(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, COLUMNS.length);
  headerRange.setValues([COLUMNS]);
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


function validateSubmissionRows(rows, dealerCode, reportDate) {
  if (!dealerCode) throw new Error('dealer_code is required');
  if (!reportDate) throw new Error('report_date is required');
  if (!Array.isArray(rows) || rows.length < MIN_EXPECTED_MODEL_ROWS) {
    throw new Error('Incomplete submission. Expected one row per model bucket.');
  }

  const seenModels = {};
  rows.forEach(row => {
    const rowDealer = String(row.dealer_code || '').trim();
    const rowDate = normaliseSheetDate(row.report_date);
    const model = String(row.model_bucket || '').trim();
    if (rowDealer !== dealerCode) throw new Error('Mixed dealer codes in one submission are not allowed');
    if (rowDate !== reportDate) throw new Error('Mixed report dates in one submission are not allowed');
    if (!model) throw new Error('model_bucket is required on every row');
    seenModels[model] = true;
  });

  if (Object.keys(seenModels).length < MIN_EXPECTED_MODEL_ROWS) {
    throw new Error('Incomplete submission. Duplicate or missing model bucket rows detected.');
  }
}

function isServerCompleteSubmission(rows, dealerCode, reportDate) {
  if (!dealerCode || !reportDate || !Array.isArray(rows) || rows.length < MIN_EXPECTED_MODEL_ROWS) return false;
  const first = rows[0] || {};
  if (!String(first.submitted_by || '').trim()) return false;
  if (!String(first.direction || '').trim()) return false;
  const models = {};
  rows.forEach(row => { if (row.model_bucket) models[String(row.model_bucket).trim()] = true; });
  return Object.keys(models).length >= MIN_EXPECTED_MODEL_ROWS;
}

function coerceDailyValue(val) {
  // Daily activity values are zero by default at submission time.
  // Forecast remains protected separately and is not coerced here.
  if (val === '' || val === null || val === undefined) return 0;
  return safeInt(val);
}

function hasExistingSubmissionRows(sheet, dealerCode, reportDate) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;
  const data = sheet.getRange(2, 1, lastRow - 1, COLUMNS.length).getValues();
  const dealerCol = COLUMNS.indexOf('dealer_code');
  const dateCol = COLUMNS.indexOf('report_date');
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][dealerCol] || '').trim() === String(dealerCode).trim()
        && normaliseSheetDate(data[i][dateCol]) === String(reportDate).trim()) return true;
  }
  return false;
}

function isSubmissionUnlocked(dealerCode, reportDate) {
  return getUnlockedDealerCodes(reportDate).indexOf(String(dealerCode || '').trim()) >= 0;
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


function getControlSheet() {
  const ss = SHEET_ID
    ? SpreadsheetApp.openById(SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName(CONTROL_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONTROL_SHEET_NAME);
    sheet.appendRow(['dealer_code', 'report_date', 'status', 'unlocked_at', 'unlocked_by']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#111111').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function recordSubmissionUnlock(dealerCode, reportDate, unlockedBy) {
  if (!dealerCode || !reportDate) throw new Error('dealer_code and report_date are required');
  const sheet = getControlSheet();
  clearSubmissionUnlock(dealerCode, reportDate);
  sheet.appendRow([String(dealerCode).trim(), String(reportDate).trim(), 'UNLOCKED', new Date().toISOString(), unlockedBy || 'OEM Dashboard']);
}

function clearSubmissionUnlock(dealerCode, reportDate) {
  const sheet = getControlSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    const rowDealer = String(data[i][0] || '').trim();
    const rowDate = normaliseSheetDate(data[i][1]);
    if (rowDealer === String(dealerCode).trim() && rowDate === String(reportDate).trim()) {
      sheet.deleteRow(i + 2);
    }
  }
}

function getUnlockedDealerCodes(reportDate) {
  const sheet = getControlSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const out = [];
  data.forEach(row => {
    const rowDealer = String(row[0] || '').trim();
    const rowDate = normaliseSheetDate(row[1]);
    const status = String(row[2] || '').trim().toUpperCase();
    if (rowDealer && rowDate === String(reportDate).trim() && status === 'UNLOCKED') out.push(rowDealer);
  });
  return Array.from(new Set(out));
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
