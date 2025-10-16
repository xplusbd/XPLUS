/**
 * Users Analytics Tracker + JSON API for "Users Dashboard" sheet
 * - Handles POST events: page_enter, page_exit
 * - Exposes GET ?format=json for dashboard and dynamic filters
 * - Auto-creates sheet/header and sets ArrayFormulas for derived columns
 *
 * Deploy as Web App (Anyone with link).
 */

const DEFAULT_SPREADSHEET_ID = '11trONTNEuQ4WfmzTynytufAsBhYzgCkCem3L5w2VUvo';
const SHEET_NAME = 'Users Dashboard';

// Column headers in exact order
const HEADERS = [
  'Timestamp',
  'Unique User ID',
  'First Visit DateTime',
  'Last Visit DateTime',
  'Last Exit DateTime',
  'Last Visited Duration (minutes)',
  'Total Visited Duration (minutes)',
  'Total Visit Count',
  'Device Type',
  'Device Model / Name',
  'Browser / WebView Info',
  'Screen Resolution / Ratio',
  'Timezone',
  'IP Address',
  'Country',
  'City',
  'Location (Latitude, Longitude)',
  'Last Visited Page / Section Name',
  'App Entry Time',
  'App Exit Time',
  'Live Status Icon',
  'User Status (Text)',
  'Days Since Last Visit',
  'Account Age (Days)',
  'Repeat Visit Type',
  'Total Sessions',
  'Active Period Range',
  'Source (optional)',
  'Notes / Admin Tag'
];

const COL = (() => {
  const map = {};
  HEADERS.forEach((h, i) => (map[h] = i + 1)); // 1-based
  return Object.freeze(map);
})();

/**
 * Web API: GET -> JSON export with filters; POST -> tracking ingest
 */
function doGet(e) {
  const q = (e && e.parameter) || {};
  const format = (q.format || '').toLowerCase();
  const sheetId = q.sheetId || DEFAULT_SPREADSHEET_ID;
  const limit = parseInt(q.limit || '', 10) || null;
  const uid = q.uid || null;
  const only = (q.only || '').toLowerCase(); // 'filters' to return only filters

  const sheet = ensureSheet(sheetId);
  ensureFormulas(sheet);

  if (format === 'json') {
    // Always read all rows once to build filters; apply limit to rows payload if requested
    const allRows = readAllRowsAsObjects(sheet, { uid: null, limit: null });
    const filters = buildFilters(allRows);

    let rows = allRows;
    if (uid) rows = rows.filter(r => String(r['Unique User ID'] || '') === String(uid));
    if (limit) rows = rows.slice(0, limit);

    if (only === 'filters') {
      return jsonResponse({ ok: true, filters, count: allRows.length });
    }

    return jsonResponse({ ok: true, rows, count: rows.length, filters });
  }

  // Simple HTML info page if format not specified
  const html = HtmlService.createHtmlOutput('<p>Users Dashboard Apps Script is running.</p>');
  return html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = parsePostBody(e);
    if (!body) return jsonResponse({ ok: false, error: 'No payload' });

    const sheetId = body.sheetId || DEFAULT_SPREADSHEET_ID;
    const sheet = ensureSheet(sheetId);
    ensureFormulas(sheet);

    const eventName = (body.event || '').toLowerCase();
    if (!body.uid) return jsonResponse({ ok: false, error: 'Missing uid' });

    let res;
    switch (eventName) {
      case 'page_enter':
        res = handlePageEnter(sheet, body);
        break;
      case 'page_exit':
        res = handlePageExit(sheet, body);
        break;
      default:
        res = { ok: false, error: 'Unknown event' };
    }
    return jsonResponse(res || { ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, error: String((err && err.message) || err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/**
 * Parse POST body from either x-www-form-urlencoded with "payload"
 * or raw JSON body.
 */
function parsePostBody(e) {
  if (!e) return null;
  // x-www-form-urlencoded: payload=<json>
  if (e.parameter && e.parameter.payload) {
    try { return JSON.parse(e.parameter.payload); } catch (_) {}
  }
  // raw JSON
  if (e.postData && e.postData.contents) {
    try { return JSON.parse(e.postData.contents); } catch (_) {}
  }
  return null;
}

/**
 * Ensure sheet exists, header row is correct
 */
function ensureSheet(spreadsheetId) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Ensure header row
  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  let needsHeader = false;
  for (let i = 0; i < HEADERS.length; i++) {
    if ((firstRow[i] || '') !== HEADERS[i]) { needsHeader = true; break; }
  }
  if (needsHeader) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  return sheet;
}

/**
 * Set/refresh ArrayFormulas for derived columns
 * W: Days Since Last Visit
 * X: Account Age (Days)
 * V: User Status (Text)
 * U: Live Status Icon
 * Y: Repeat Visit Type
 * Z: Total Sessions
 * AA: Active Period Range
 */
function ensureFormulas(sheet) {
  // W2: Days Since Last Visit = TODAY() - DATEVALUE(D)
  const w2 = '=ARRAYFORMULA(IF(ROW(D2:D)=2, IF(D2<>"","",), IF(D2:D="", "", TODAY() - DATEVALUE(D2:D))))';
  setIfDifferent(sheet, 2, COL['Days Since Last Visit'], w2);

  // X2: Account Age = TODAY() - DATEVALUE(C)
  const x2 = '=ARRAYFORMULA(IF(ROW(C2:C)=2, IF(C2<>"","",), IF(C2:C="", "", TODAY() - DATEVALUE(C2:C))))';
  setIfDifferent(sheet, 2, COL['Account Age (Days)'], x2);

  // V2: User Status (New if age<=15; Active if W<=30; Inactive if W>30)
  const v2 = [
    '=ARRAYFORMULA(IF(LEN(B2:B)=0, "",',
    ' IF(X2:X<=15, "New",',
    '  IF(W2:W>30, "Inactive", "Active")',
    ' )',
    '))'
  ].join('');
  setIfDifferent(sheet, 2, COL['User Status (Text)'], v2);

  // U2: Live Status Icon
  // 🟢 if entry set, exit empty, and NOW()-Entry <= 0.1h (6m)
  // else 🔵 if W<=1, else 🟣 for Active, 🔴 for Inactive, 🆕 for New
  const u2 = [
    '=ARRAYFORMULA(IF(LEN(B2:B)=0, "",',
    ' IF( (LEN(S2:S)>0) * (LEN(T2:T)=0) * ((NOW()-S2:S) <= 0.1/24), "🟢",',
    '  IF(W2:W<=1, "🔵",',
    '    IF(V2:V="Inactive", "🔴",',
    '      IF(V2:V="New", "🆕", "🟣")',
    '    )',
    '  )',
    ' )',
    '))'
  ].join('');
  setIfDifferent(sheet, 2, COL['Live Status Icon'], u2);

  // Y2: Repeat Visit Type (First / Returning based on Total Visit Count)
  const y2 = '=ARRAYFORMULA(IF(LEN(B2:B)=0, "", IF(H2:H<=1, "First", "Returning")))';
  setIfDifferent(sheet, 2, COL['Repeat Visit Type'], y2);

  // Z2: Total Sessions = Total Visit Count
  const z2 = '=ARRAYFORMULA(IF(LEN(B2:B)=0, "", H2:H))';
  setIfDifferent(sheet, 2, COL['Total Sessions'], z2);

  // AA2: Active Period Range from W (≤24h, ≤7d, ≤30d, >30d)
  const aa2 = [
    '=ARRAYFORMULA(IF(LEN(B2:B)=0, "",',
    ' IF(W2:W<=1,"≤24h",',
    '  IF(W2:W<=7,"≤7d",',
    '   IF(W2:W<=30,"≤30d",">30d")',
    '  )',
    ' )',
    '))'
  ].join('');
  setIfDifferent(sheet, 2, COL['Active Period Range'], aa2);
}

function setIfDifferent(sheet, row, col, formula) {
  const cell = sheet.getRange(row, col);
  if (cell.getFormula() !== formula) {
    cell.setFormula(formula);
  }
}

/**
 * Handle page_enter event
 */
function handlePageEnter(sheet, body) {
  const uid = String(body.uid || '').trim();
  const now = new Date();
  const entryIso = body.entryTime || now.toISOString();

  const rowIndex = findRowByUID(sheet, uid);
  if (rowIndex === -1) {
    // New user row
    const row = new Array(HEADERS.length).fill('');
    row[idx('Timestamp')] = now;
    row[idx('Unique User ID')] = uid;
    row[idx('First Visit DateTime')] = body.firstVisitIso ? new Date(body.firstVisitIso) : new Date(entryIso);
    row[idx('Last Visit DateTime')] = new Date(entryIso);
    row[idx('Last Exit DateTime')] = '';
    row[idx('Last Visited Duration (minutes)')] = 0;
    row[idx('Total Visited Duration (minutes)')] = 0;
    row[idx('Total Visit Count')] = 1;
    row[idx('Device Type')] = body.deviceType || '';
    row[idx('Device Model / Name')] = body.deviceModel || '';
    row[idx('Browser / WebView Info')] = body.userAgent || '';
    row[idx('Screen Resolution / Ratio')] = body.screenResolution || '';
    row[idx('Timezone')] = body.timezone || '';
    row[idx('IP Address')] = body.ip || '';
    row[idx('Country')] = body.country || '';
    row[idx('City')] = body.city || '';
    row[idx('Location (Latitude, Longitude)')] = body.location || '';
    row[idx('Last Visited Page / Section Name')] = body.sectionName || '';
    row[idx('App Entry Time')] = new Date(entryIso);
    row[idx('App Exit Time')] = '';
    // U,V,W,X via formulas
    row[idx('Repeat Visit Type')] = ''; // via formula
    row[idx('Total Sessions')] = '';    // via formula
    row[idx('Active Period Range')] = ''; // via formula
    row[idx('Source (optional)')] = body.source || '';
    row[idx('Notes / Admin Tag')] = '';

    sheet.appendRow(row);
    return { ok: true, created: true, uid };
  } else {
    // Update existing
    const rng = sheet.getRange(rowIndex, 1, 1, HEADERS.length);
    const vals = rng.getValues()[0];

    const prevTotalDur = toNumber(vals[idx('Total Visited Duration (minutes)')]);
    const prevVisitCount = toNumber(vals[idx('Total Visit Count')]);

    vals[idx('Timestamp')] = now;
    vals[idx('Last Visit DateTime')] = new Date(entryIso);
    vals[idx('App Entry Time')] = new Date(entryIso);
    vals[idx('App Exit Time')] = '';
    vals[idx('Last Visited Page / Section Name')] = body.sectionName || vals[idx('Last Visited Page / Section Name')] || '';

    // Update metadata (latest known)
    vals[idx('Device Type')] = body.deviceType || vals[idx('Device Type')] || '';
    vals[idx('Device Model / Name')] = body.deviceModel || vals[idx('Device Model / Name')] || '';
    vals[idx('Browser / WebView Info')] = body.userAgent || vals[idx('Browser / WebView Info')] || '';
    vals[idx('Screen Resolution / Ratio')] = body.screenResolution || vals[idx('Screen Resolution / Ratio')] || '';
    vals[idx('Timezone')] = body.timezone || vals[idx('Timezone')] || '';
    vals[idx('IP Address')] = body.ip || vals[idx('IP Address')] || '';
    vals[idx('Country')] = body.country || vals[idx('Country')] || '';
    vals[idx('City')] = body.city || vals[idx('City')] || '';
    vals[idx('Location (Latitude, Longitude)')] = body.location || vals[idx('Location (Latitude, Longitude)')] || '';
    vals[idx('Source (optional)')] = body.source || vals[idx('Source (optional)')] || '';

    // Counters
    vals[idx('Total Visit Count')] = prevVisitCount + 1;
    vals[idx('Last Visited Duration (minutes)')] = toNumber(vals[idx('Last Visited Duration (minutes)')]) || 0;
    vals[idx('Total Visited Duration (minutes)')] = prevTotalDur || 0;

    rng.setValues([vals]);
    return { ok: true, updated: true, uid };
  }
}

/**
 * Handle page_exit event
 */
function handlePageExit(sheet, body) {
  const uid = String(body.uid || '').trim();
  const now = new Date();
  const exitIso = body.exitTime || now.toISOString();
  const durMin = isFinite(body.durationMinutes) ? Number(body.durationMinutes) : 0;

  let rowIndex = findRowByUID(sheet, uid);
  if (rowIndex === -1) {
    // If no existing row, create minimal and set exit
    const row = new Array(HEADERS.length).fill('');
    row[idx('Timestamp')] = now;
    row[idx('Unique User ID')] = uid;
    row[idx('First Visit DateTime')] = '';
    row[idx('Last Visit DateTime')] = '';
    row[idx('Last Exit DateTime')] = new Date(exitIso);
    row[idx('Last Visited Duration (minutes)')] = durMin;
    row[idx('Total Visited Duration (minutes)')] = durMin;
    row[idx('Total Visit Count')] = 1;
    row[idx('App Entry Time')] = '';
    row[idx('App Exit Time')] = new Date(exitIso);
    sheet.appendRow(row);
    return { ok: true, created: true, uid, note: 'Created on exit' };
  } else {
    const rng = sheet.getRange(rowIndex, 1, 1, HEADERS.length);
    const vals = rng.getValues()[0];

    const prevTotalDur = toNumber(vals[idx('Total Visited Duration (minutes)')]) || 0;

    vals[idx('Timestamp')] = now;
    vals[idx('Last Exit DateTime')] = new Date(exitIso);
    vals[idx('App Exit Time')] = new Date(exitIso);
    vals[idx('Last Visited Duration (minutes)')] = durMin;
    vals[idx('Total Visited Duration (minutes)')] = round2(prevTotalDur + durMin);

    rng.setValues([vals]);
    return { ok: true, updated: true, uid };
  }
}

/**
 * Export rows as objects for dashboard
 */
function readAllRowsAsObjects(sheet, opts) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const lastCol = HEADERS.length;

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rows = [];
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const uid = String(r[idx('Unique User ID')] || '').trim();
    if (!uid) continue;

    if (opts && opts.uid && uid !== opts.uid) continue;

    const obj = {};
    for (let c = 0; c < HEADERS.length; c++) {
      const header = HEADERS[c];
      let val = r[c];
      if (val instanceof Date) {
        // convert to ISO string
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX");
      }
      obj[header] = val;
    }
    rows.push(obj);
    if (opts && opts.limit && rows.length >= opts.limit) break;
  }
  return rows;
}

/**
 * Build filters payload (distinct lists) from rows
 */
function buildFilters(rows) {
  const uniq = (arr) => Array.from(new Set((arr || []).filter(v => v != null && String(v).trim() !== '')));

  const statusValues = uniq(rows.map(r => String(r['User Status (Text)'] || '')));
  const repeatValues = uniq(rows.map(r => String(r['Repeat Visit Type'] || '')));
  // Include 'Returning' in status if present in repeat values
  const status = uniq(statusValues.concat(repeatValues.indexOf('Returning') >= 0 ? ['Returning'] : []));

  const country = uniq(rows.map(r => r['Country'])).sort();
  const city = uniq(rows.map(r => r['City'])).sort();
  const ip = uniq(rows.map(r => r['IP Address']));

  const timestampRaw = uniq(rows.map(r => r['Timestamp']));
  // Sort timestamps desc
  const timestamp = timestampRaw.sort(function(a,b){
    var da = new Date(a).getTime();
    var db = new Date(b).getTime();
    if (isNaN(da) && isNaN(db)) return 0;
    if (isNaN(da)) return 1;
    if (isNaN(db)) return -1;
    return db - da;
  }).slice(0, 200); // limit to first 200 distinct timestamps

  const timeRange = [
    'All',
    'Today',
    'Last 7 Days',
    'Last 15 Days',
    'Last 30 Days',
    'Last 3 Months',
    'Last 6 Months',
    'Last 1 Year'
  ];

  return { timeRange, status, country, city, ip, timestamp };
}

/** Helpers */
function idx(header) { return COL[header] - 1; } // 0-based for row arrays
function toNumber(v) {
  if (v === '' || v == null) return 0;
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}
function round2(n) { return Math.round(n * 100) / 100; }

function findRowByUID(sheet, uid) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const uidCol = COL['Unique User ID'];
  const range = sheet.getRange(2, uidCol, lastRow - 1, 1);
  const vals = range.getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || '').trim() === uid) {
      return i + 2; // row index (1-based)
    }
  }
  return -1;
}

function jsonResponse(obj) {
  const out = ContentService.createTextOutput(JSON.stringify(obj));
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}
