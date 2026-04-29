/* ============================================================
   GWM Australia – Daily Sales Reporting Tool
   Shared Data & Utilities  (app.js)
   ============================================================ */

'use strict';

/* ── Configuration ──────────────────────────────────────── */
const GWM_CONFIG = {
  // Paste your deployed Google Apps Script Web App URL here.
  // Leave empty to use demo / offline mode.
  scriptUrl: 'https://script.google.com/macros/s/AKfycbxh7-IUcOe6tGrkCqK1sYrzb4LBhPoiwwePqkEd_sab2vAUGihrEGIer2Gm5VlUFH14sA/exec',
  region: 'Southern Region',
  cutoffHour: 10,   // Submissions after 10:00am are flagged late

  // First date the system should enforce from. Set this to your go-live date.
  // Leave blank to enforce from the first day of the current month.
  reportingStartDate: '',

  // Sunday reporting is opt-in by dealer. Most dealers will not be forced to report Sundays.
  // Add dealer codes here only if they trade Sundays and need Sunday activity captured.
  sundayTradingDealers: [],
};

/* ── Dealer list ─────────────────────────────────────────── */
const DEALERS = [
  { code: 'H3100', name: 'Berwick GWM' },
  { code: 'H3101', name: 'Doncaster GWM' },
  { code: 'H3104', name: 'Astoria GWM' },
  { code: 'H3107', name: 'Melton GWM' },
  { code: 'H3128', name: 'Werribee GWM' },
  { code: 'H3161', name: 'Peninsula GWM' },
  { code: 'H3163', name: 'Ringwood GWM' },
  { code: 'H3182', name: 'Ralph D\'Silva GWM' },
  { code: 'H3167', name: 'Essendon GWM' },
  { code: 'H3318', name: 'Thompson GWM (Shepparton)' },
  { code: 'H3315', name: 'Valley GWM' },
  { code: 'H7215', name: 'Hobart GWM' },
  { code: 'H7309', name: 'Launceston GWM' },
  { code: 'H3230', name: 'Geelong GWM' },
  { code: 'H3236', name: 'Bendigo GWM' },
  { code: 'H3176', name: 'South Morang GWM' },
  { code: 'H3179', name: 'Burwood GWM' },
  { code: 'H3185', name: 'Knox GWM' },
  { code: 'H3196', name: 'Western GWM' },
  { code: 'H3195', name: 'Pakenham GWM' },
  { code: 'H3239', name: 'Ballarat GWM' },
  { code: 'H3336', name: 'Thompson GWM (Echuca)' },
  { code: 'H3342', name: 'Horsham GWM' },
  { code: 'H3188', name: 'Blackburn GWM' },
  { code: 'H3191', name: 'Lilydale GWM' },
  { code: 'H3333', name: 'Blacklocks GWM (Albury)' },
  { code: 'H3330', name: 'Warrnambool GWM' },
  { code: 'H3110', name: 'Dandenong GWM' },
  { code: 'H3345', name: 'Peter Dullard GWM (Bairnsdale)' },
  { code: 'H3111', name: 'Melbourne CBD' },
];

/* ── Model buckets ───────────────────────────────────────── */
const MODEL_BUCKETS = [
  'Jolion ICE',
  'Jolion HEV',
  'H6 Petrol',
  'H6 HEV',
  'H6 PHEV',
  'H6 GT Petrol',
  'H6 GT PHEV',
  'H7 HEV',
  'Tank 300 Petrol',
  'Tank 300 Diesel',
  'Tank 300 PHEV',
  'Tank 500 Hybrid',
  'Tank 500 PHEV',
  'Tank 500 Diesel',
  'Cannon Diesel',
  'Cannon Alpha Diesel',
  'Cannon Alpha PHEV',
  'Ora',
  'Ora 5',
];

/* ── Grid column definitions ─────────────────────────────── */
const GRID_COLS = [
  { key: 'enquiry',     label: 'Enquiry',    short: 'ENQ' },
  { key: 'test_drives', label: 'Test Drives', short: 'TD' },
  { key: 'new_sold',    label: 'New Sold',   short: 'SOLD' },
  { key: 'fleet',       label: 'Fleet (5+)', short: 'FLT' },
  { key: 'demo_sold',   label: 'Demo Sold',  short: 'DEMO' },
  { key: 'forecast',    label: 'Forecast',   short: 'FCST' },
];

/* ── Date helpers ────────────────────────────────────────── */
function todayISO() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

function nowTimestamp() {
  return new Date().toISOString();
}

/* Reporting day helper
   Reporting is always retrospective: today's input is for yesterday's dealer activity.
   Do not use a time-based rollover. The day has to be complete before it is reported.
*/
function reportingDateISO() {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  return toISODate(d);
}

function toISODate(d) {
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

function parseISODate(iso) {
  const [y, m, d] = String(iso || '').split('-').map(Number);
  return new Date(y, (m || 1) - 1, d || 1);
}

function isoOffset(iso, days) {
  const d = parseISODate(iso);
  d.setDate(d.getDate() + days);
  return toISODate(d);
}

function isSundayISO(iso) {
  return parseISODate(iso).getDay() === 0;
}

function isSundayRequiredForDealer(dealerCode) {
  return (GWM_CONFIG.sundayTradingDealers || []).map(String).includes(String(dealerCode || '').trim());
}

function reportingStartDateISO() {
  if (GWM_CONFIG.reportingStartDate) return GWM_CONFIG.reportingStartDate;
  const d = new Date();
  d.setDate(1);
  return toISODate(d);
}

function buildRequiredReportDates(dealerCode, startISO, endISO) {
  const dates = [];
  let cur = startISO || reportingStartDateISO();
  const end = endISO || reportingDateISO();
  while (cur <= end) {
    if (!isSundayISO(cur) || isSundayRequiredForDealer(dealerCode)) dates.push(cur);
    cur = isoOffset(cur, 1);
  }
  return dates;
}

function firstMissingReportDate(dealerCode, submittedDates, startISO, endISO) {
  const submitted = new Set((submittedDates || []).map(String));
  return buildRequiredReportDates(dealerCode, startISO, endISO).find(d => !submitted.has(d)) || '';
}

function monthKeyFromISO(iso) {
  return String(iso || todayISO()).slice(0, 7);
}

function formatDate(iso) {
  if (!iso) return '';
  const [y, m, d] = iso.split('-');
  return `${d}/${m}/${y}`;
}

function isLate() {
  return new Date().getHours() >= GWM_CONFIG.cutoffHour;
}

function formatTime(isoString) {
  if (!isoString) return '';
  const d = new Date(isoString);
  return d.toLocaleTimeString('en-AU', { hour: '2-digit', minute: '2-digit', hour12: true });
}

/* ── Number helpers ──────────────────────────────────────── */
function safeInt(val) {
  const n = parseInt(val, 10);
  return isNaN(n) || n < 0 ? 0 : n;
}

/* ── Toast notifications ─────────────────────────────────── */
function showToast(msg, type = 'default', duration = 3500) {
  let toast = document.getElementById('gwm-toast');
  if (!toast) {
    toast = document.createElement('div');
    toast.id = 'gwm-toast';
    toast.className = 'toast';
    document.body.appendChild(toast);
  }
  toast.textContent = msg;
  toast.className = `toast ${type}`;
  requestAnimationFrame(() => {
    requestAnimationFrame(() => toast.classList.add('show'));
  });
  clearTimeout(toast._timer);
  toast._timer = setTimeout(() => toast.classList.remove('show'), duration);
}

/* ── Submission confirmation helper ───────────────────── */
async function waitForSubmissionConfirmation(dealerCode, reportDate, attempts = 8) {
  for (let i = 0; i < attempts; i++) {
    await new Promise(r => setTimeout(r, i < 2 ? 900 : 1400));
    try {
      const state = await fetchRows({ date: reportDate, dealer: dealerCode, include_controls: '1' });
      const rows = Array.isArray(state) ? state : (state?.rows || []);
      const unlocked = Array.isArray(state?.unlockedDealerCodes) ? state.unlockedDealerCodes.map(x => String(x).trim()) : [];
      if (rows.length >= MODEL_BUCKETS.length && !unlocked.includes(String(dealerCode).trim())) return rows;
    } catch (err) {
      console.warn('Submission confirmation check failed:', err);
    }
  }
  throw new Error('Submission was sent, but the backend did not confirm the saved rows. Refresh and check the dashboard before submitting again.');
}

/* ── API call (POST to Apps Script) ─────────────────────── */
async function postRows(rows) {
  const url = GWM_CONFIG.scriptUrl;
  if (!url) throw new Error('NO_URL');
  if (!Array.isArray(rows) || !rows.length) throw new Error('No rows to submit');

  // Preflight lock check. Apps Script POST uses no-cors, so this GET check gives the UI
  // a readable control-layer response before we send anything.
  const first = rows[0];
  const dealerCode = String(first.dealer_code || '').trim();
  const reportDate = String(first.report_date || '').trim();
  if (dealerCode && reportDate) {
    if (reportDate >= todayISO()) {
      throw new Error("Cannot submit today's date. Reporting is always for completed prior days only.");
    }
    const state = await fetchRows({ date: reportDate, dealer: dealerCode, include_controls: '1' });
    const existingRows = Array.isArray(state) ? state : (state?.rows || []);
    const unlocked = Array.isArray(state?.unlockedDealerCodes) ? state.unlockedDealerCodes.map(x => String(x).trim()) : [];
    if (existingRows.length && !unlocked.includes(dealerCode)) {
      throw new Error('This dealer/date is locked. Reopen it from the dashboard before resubmitting.');
    }
  }

  await fetch(url, {
    method: 'POST',
    mode: 'no-cors',               // Apps Script requires no-cors
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ rows }),
  });

  // no-cors POSTs are opaque. Confirm via readable GET before the UI marks success.
  await waitForSubmissionConfirmation(dealerCode, reportDate);
  return { ok: true, confirmed: true };
}

/* ── API control call (POST to Apps Script) ─────────────── */
async function postControl(payload) {
  const url = GWM_CONFIG.scriptUrl;
  if (!url) throw new Error('NO_URL');

  await fetch(url, {
    method: 'POST',
    mode: 'no-cors',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  return { ok: true };
}

/* ── API call (GET from Apps Script) ────────────────────── */
async function fetchRows(params = {}) {
  const url = GWM_CONFIG.scriptUrl;
  if (!url) return null;
  const qs = new URLSearchParams(params).toString();
  const res = await fetch(`${url}?${qs}`);
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  return res.json();
}

/* ── Demo data generator (used when no scriptUrl set) ───── */
function generateDemoData(dateISO) {
  const rows = [];
  const submitDealers = DEALERS.slice(0, 22);
  submitDealers.forEach(dealer => {
    const late = Math.random() < 0.25;
    const ts = new Date(dateISO + 'T' + (late ? '10' : '08') + ':' + String(Math.floor(Math.random()*59)).padStart(2,'0') + ':00');
    MODEL_BUCKETS.forEach(model => {
      const active = Math.random() < 0.6;
      rows.push({
        submitted_at: ts.toISOString(),
        report_date: dateISO,
        is_late: late,
        dealer_code: dealer.code,
        dealer_name: dealer.name,
        region: GWM_CONFIG.region,
        submitted_by: 'Demo User',
        direction: ['Up','Flat','Down'][Math.floor(Math.random()*3)],
        is_complete_submission: true,
        input_method: 'demo',
        submission_duration_seconds: 45,
        last_updated_at: ts.toISOString(),
        model_bucket: model,
        enquiry:     active ? Math.floor(Math.random()*8) : 0,
        test_drives: active ? Math.floor(Math.random()*4) : 0,
        new_sold:    active ? Math.floor(Math.random()*3) : 0,
        fleet:       Math.random() < 0.1 ? Math.floor(Math.random()*10)+5 : 0,
        demo_sold:   Math.random() < 0.15 ? 1 : 0,
        forecast:    active ? Math.floor(Math.random()*5) : 0,
      });
    });
  });
  return rows;
}

/* ── CSV export ──────────────────────────────────────────── */
function exportCSV(rows, filename) {
  const headers = [
    'submitted_at','report_date','is_late','dealer_code','dealer_name',
    'region','submitted_by','direction','is_complete_submission','input_method',
    'submission_duration_seconds','last_updated_at','model_bucket',
    'enquiry','test_drives','new_sold','fleet_5_plus','demo_sold','forecast'
  ];
  const escape = v => {
    const s = String(v ?? '');
    return s.includes(',') || s.includes('"') || s.includes('\n')
      ? `"${s.replace(/"/g,'""')}"` : s;
  };
  const lines = [headers.join(',')];
  rows.forEach(r => {
    lines.push([
      r.submitted_at, r.report_date, r.is_late ? 'TRUE' : 'FALSE',
      r.dealer_code, r.dealer_name, r.region, r.submitted_by, r.direction,
      r.is_complete_submission ? 'TRUE' : 'FALSE',
      r.input_method, r.submission_duration_seconds, r.last_updated_at,
      r.model_bucket, r.enquiry, r.test_drives, r.new_sold,
      r.fleet, r.demo_sold, r.forecast,
    ].map(escape).join(','));
  });
  const blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = filename || `gwm_report_${todayISO()}.csv`;
  link.click();
}
