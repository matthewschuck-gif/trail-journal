/**
 * Trail Journal -- Apps Script backend
 * File 3 of N: doPost/doGet router + generic CRUD over the Sheet tabs.
 *
 * DESIGN NOTE ON CORS (read this before wiring up the frontend in file 9):
 * Apps Script web apps don't let you set custom response headers, and they don't handle
 * CORS preflight (OPTIONS) requests. The fix -- and the reason frontend fetch() calls need
 * a small tweak -- is to send requests with `Content-Type: text/plain;charset=utf-8` instead
 * of `application/json`. That keeps the browser from sending a preflight at all, and Google's
 * infrastructure serves script.google.com/exec responses cross-origin without extra headers.
 * The body is still just JSON.stringify(...) text; only the header name changes. doPost()
 * below parses e.postData.contents as JSON regardless of what Content-Type was declared.
 *
 * DESIGN NOTE ON SECURITY (equivalent, not identical, to the original):
 * The original app protected writes with a Supabase anon key + Row Level Security. An anon
 * key embedded in frontend JS is not a secret -- anyone can read it from page source -- so
 * the *real* protection was RLS running server-side in Postgres. Apps Script has no equivalent
 * of RLS. What this router does instead: every request must include a shared APP_TOKEN (set in
 * Script Properties, also embedded as a constant in the frontend JS -- same trust model as the
 * anon key had). This filters out random internet scanners hitting the URL, but -- exactly like
 * the anon key -- it is visible to anyone who views page source. It is not real authentication.
 * The one place this matters most is incident_reports (real student names). Treat that tab's
 * sharing settings in the actual Google Sheet as the real access control, same as you would have
 * needed to double check RLS policies on that table in Supabase.
 */

function doPost(e) {
  return handleRequest_(e, 'POST');
}

function doGet(e) {
  return handleRequest_(e, 'GET');
}

function handleRequest_(e, method) {
  try {
    let body;
    if (method === 'POST') {
      body = JSON.parse(e.postData.contents);
    } else {
      // GET: everything comes through query params. filters/row/patch arrive as a JSON string.
      body = {
        action: e.parameter.action,
        table: e.parameter.table,
        filters: e.parameter.filters ? JSON.parse(e.parameter.filters) : {},
        limit: e.parameter.limit ? Number(e.parameter.limit) : undefined,
        token: e.parameter.token,
      };
    }

    checkAppToken_(body.token);
    enforceAdminGate_(body); // throws if this action needs staff Google Sign-In and none was provided (see 05_AdminAuth.gs)

    let result;
    switch (body.action) {
      case 'insert':
        result = genericInsert_(body.table, body.row);
        break;
      case 'update':
        result = genericUpdate_(body.table, body.match, body.patch);
        break;
      case 'query':
        result = genericQuery_(body.table, body.filters || {}, body.limit);
        break;
      case 'ai':
        result = callAiProxy_(body.type, body.payload);
        break;
      case 'checkStaffPassword':
        result = checkStaffPassword_(body.password);
        break;
      case 'sendFollowupEmails':
        result = sendFollowupEmails_(body.payload);
        break;
      case 'sendSummaryEmail':
        result = sendSummaryEmail_(body.payload);
        break;
      case 'sendOfficeRecordEmail':
        result = sendOfficeRecordEmail_(body.payload);
        break;
      case 'uploadPeaksPoster':
        result = uploadPeaksPoster_(body.payload);
        break;
      case 'resetPeaksPoster':
        result = resetPeaksPoster_(body.payload);
        break;
      case 'setOfficeRecordRecipients':
        result = setOfficeRecordRecipients_(body.payload);
        break;
      default:
        throw new Error('Unknown action: ' + body.action);
    }

    return jsonOut_({ ok: true, data: result });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function checkAppToken_(token) {
  const expected = getProp_('APP_TOKEN');
  if (token !== expected) throw new Error('Invalid or missing token');
}

// --- JSON-valued columns per table (Postgres jsonb/array -> stored as JSON text in the cell) ---
const JSON_COLUMNS = {
  reflections: ['part1_json', 'part2_json', 'part3_json', 'ai_staff_summary', 'ai_reading_list', 'conference_brief'],
  responder_reflections: ['emotions_json'],
  panel_sessions: ['panel_questions', 'impact_statements', 'recommendation_categories', 'followup_dates', 'followup_notes', 'ai_recommendation'],
  followup_records: ['attendee_emails', 'calendar_status'],
  incident_reports: [],
  site_config: [],
};

// --- Default values applied on insert, mirroring the Postgres column defaults ---
const DEFAULTS = {
  reflections: { status: 'submitted', has_flags: false, prior_count: 0, panel_completed: false },
  responder_reflections: { status: 'submitted' },
  panel_sessions: { status: 'open', clash_point_awarded: false },
  incident_reports: { status: 'open' },
  followup_records: { calendar_status: { day2: 'pending', week2: 'pending', month2: 'pending' } },
  site_config: {},
};

function genericInsert_(table, row) {
  const sheet = getSheet_(table);
  const headers = SCHEMA[table];
  const jsonCols = JSON_COLUMNS[table] || [];
  const defaults = DEFAULTS[table] || {};

  const full = Object.assign({}, defaults, row || {});
  if (headers.indexOf('id') !== -1 && !full.id) full.id = newUuid_();
  if (headers.indexOf('created_at') !== -1 && !full.created_at) full.created_at = nowIso_();

  const rowValues = headers.map(function (col) {
    let v = full[col];
    if (v === undefined || v === null) return '';
    if (jsonCols.indexOf(col) !== -1 && typeof v !== 'string') return JSON.stringify(v);
    return v;
  });

  sheet.appendRow(rowValues);
  return deserializeRow_(table, full);
}

function genericUpdate_(table, match, patch) {
  const sheet = getSheet_(table);
  const headers = SCHEMA[table];
  const jsonCols = JSON_COLUMNS[table] || [];
  const data = sheet.getDataRange().getValues();

  const matchCols = Object.keys(match || {});
  if (matchCols.length === 0) throw new Error('update requires at least one match column');

  let updatedCount = 0;
  for (let i = 1; i < data.length; i++) {
    const isMatch = matchCols.every(function (col) {
      const colIdx = headers.indexOf(col);
      return String(data[i][colIdx]) === String(match[col]);
    });
    if (!isMatch) continue;

    const rowNum = i + 1;
    Object.keys(patch || {}).forEach(function (col) {
      const colIdx = headers.indexOf(col);
      if (colIdx === -1) return;
      let v = patch[col];
      if (jsonCols.indexOf(col) !== -1 && typeof v !== 'string') v = JSON.stringify(v);
      sheet.getRange(rowNum, colIdx + 1).setValue(v === undefined || v === null ? '' : v);
    });
    updatedCount++;
  }
  if (updatedCount === 0) throw new Error('No matching row found in ' + table + ' for ' + JSON.stringify(match));
  return { updated: updatedCount };
}

function genericQuery_(table, filters, limit) {
  const sheet = getSheet_(table);
  const headers = SCHEMA[table];
  const data = sheet.getDataRange().getValues();
  const filterCols = Object.keys(filters || {});

  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const matches = filterCols.every(function (col) {
      const colIdx = headers.indexOf(col);
      if (colIdx === -1) return true;
      return String(data[i][colIdx]) === String(filters[col]);
    });
    if (!matches) continue;

    const obj = {};
    headers.forEach(function (col, idx) { obj[col] = data[i][idx]; });
    rows.push(deserializeRow_(table, obj));
    if (limit && rows.length >= limit) break;
  }
  return rows;
}

function deserializeRow_(table, obj) {
  const jsonCols = JSON_COLUMNS[table] || [];
  const out = Object.assign({}, obj);
  jsonCols.forEach(function (col) {
    if (typeof out[col] === 'string' && out[col] !== '') {
      try { out[col] = JSON.parse(out[col]); } catch (e) { /* leave as string if not parseable */ }
    }
  });
  return out;
}
