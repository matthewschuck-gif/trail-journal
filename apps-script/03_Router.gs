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
  // No ?action= → one of the pages (08_Web.gs). ?action= is the old token API.
  if (!e || !e.parameter || !e.parameter.action) return serveTrailPage_(e);
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

    delete body._staffEmail; // only tjCall (08_Web.gs) sets this, from the Google sign-in
    delete body._caller;
    // Other EMS tools (Accountability system, Climb-er) reach two Trailback actions with their
    // own key (Script Property TRAILBACK_API_KEY), even after LEGACY_API is OFF. 12_TrailbackCases.gs
    if (TRAILBACK_API_ACTIONS_.indexOf(body.action) !== -1 && body.key) {
      return jsonOut_({ ok: true, data: trailbackApi_(body) });
    }
    checkAppToken_(body.token);
    // After the switch to the Apps Script pages (Settings LEGACY_API = OFF), the old token route
    // only takes the exit auto-save beacon a closing journal page sends.
    if (String(webSetting_('LEGACY_API')).toUpperCase() === 'OFF' && body.action !== 'exitAutoSave') {
      throw new Error('This address is retired. Open Trail Journal at trailjournal.org.');
    }
    return jsonOut_({ ok: true, data: dispatch_(body) });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

// Runs one request (after the token or sign-in check). Shared by doPost and tjCall.
function dispatch_(body) {
    enforceAdminGate_(body); // throws if this action needs staff access and none was provided (see 05_AdminAuth.gs)

    // A panel's approval is recorded only by approvePanelPlan_ (11_Trailback.gs), from the
    // Google sign-in. The generic writes can't set or change it.
    if (body.table === TABS.PANEL_SESSIONS) {
      if (body.action === 'insert') body.row = stripPanelApproval_(body.row);
      if (body.action === 'update') body.patch = stripPanelApproval_(body.patch);
    }

    let result;
    switch (body.action) {
      case 'followupDates': // 3 / 10 / 30 school days after a date (11_Trailback.gs)
        result = followupSchedule_(body.payload);
        break;
      case 'approvePanelPlan':
        result = approvePanelPlan_(body);
        break;
      case 'submitPanelSession':
        result = submitPanelSession_(body);
        break;
      case 'clashRequests':
        result = clashRequests_();
        break;
      case 'decideClashRequest':
        result = decideClashRequest_(body);
        break;
      case 'trailbackCases': // 12_TrailbackCases.gs
        result = trailbackCases_(body);
        break;
      case 'trailbackCase':
        result = trailbackCase_(body);
        break;
      case 'updateTrailbackCase':
        result = updateTrailbackCase_(body);
        break;
      case 'createTrailbackCase':
        result = createTrailbackCase_(body);
        break;
      case 'trailbackCasesByStudent':
        result = trailbackCasesByStudent_(body);
        break;
      case 'insert':
        result = genericInsert_(body.table, body.row);
        break;
      case 'update':
        result = genericUpdate_(body.table, body.match, body.patch);
        break;
      case 'query':
        result = isStudentHistoryQuery_(body) && !hasStaffPassword_(body)
          ? studentHistory_(body.filters)
          : genericQuery_(body.table, body.filters || {}, body.limit);
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
      case 'exitAutoSave':
        result = exitAutoSave_(body.payload);
        break;
      case 'upsertReflection': // every save from the journal page, by session (10_Records.gs)
        result = upsertReflection_(body.row);
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
      case 'clashStandings':
        result = clashStandings_();
        break;
      // The panel's questions and the repair plan check-ins (14_RepairPlan.gs). Staff sign-in only.
      // Action ideas, the Give / Get board (15_GiveGet.gs) and Summit celebrations (16_Summit.gs).
      // The family consent page (17_Pipeline.gs): public, by its one-time link.
      case 'trailbackOutcomes': // 18_Outcomes.gs (leads + coordinator)
        result = trailbackOutcomes_(body);
        break;
      case 'consentLookup':
        result = consentLookup_(body);
        break;
      case 'consentAnswer':
        result = consentAnswer_(body);
        break;
      case 'actionIdeas':
        result = actionIdeas_(body);
        break;
      case 'giveGetBoard':
        result = giveGetBoard_(body);
        break;
      case 'giveGetSave':
        result = giveGetSave_(body);
        break;
      case 'summitList':
        result = summitList_(body);
        break;
      case 'summitSave':
        result = summitSave_(body);
        break;
      case 'panelQuestions':
        result = panelQuestions_(body);
        break;
      case 'repairCheckins':
        result = repairCheckins_(body);
        break;
      case 'repairCheckin':
        result = repairCheckin_(body);
        break;
      // Trailback Panel sign-up (13_PanelSignup.gs). Each checks the signed-in staff account itself.
      case 'panelBoard':
        result = panelBoard_(body);
        break;
      case 'panelSaveCase':
        result = panelSaveCase_(body);
        break;
      case 'panelSignUp':
        result = panelSignUp_(body);
        break;
      case 'panelWithdraw':
        result = panelWithdraw_(body);
        break;
      case 'panelSendBrief':
        result = panelSendBrief_(body);
        break;
      case 'awardClashPoint':
        result = awardClashPoint_(body.payload);
        break;
      case 'overrideInsightLevelHigh':
        result = overrideInsightLevelHigh_(body.payload);
        break;
      default:
        throw new Error('Unknown action: ' + body.action);
    }
    return result;
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
  panel_sessions: ['panel_questions', 'impact_statements', 'recommendation_categories', 'followup_dates', 'followup_notes', 'ai_recommendation', 'plan_actions', 'panel_strengths', 'stage_minutes'],
  followup_records: ['attendee_emails', 'calendar_status'],
  incident_reports: [],
  site_config: [],
  clash_requests: [],
  panel_cases: ['checkin_notes', 'history', 'referral_json'],
};

// --- Default values applied on insert, mirroring the Postgres column defaults ---
const DEFAULTS = {
  reflections: { status: 'submitted', has_flags: false, prior_count: 0, panel_completed: false, clash_flagged: false, clash_points_awarded: false },
  responder_reflections: { status: 'submitted' },
  panel_sessions: { status: 'open', clash_point_awarded: false },
  incident_reports: { status: 'open' },
  followup_records: { calendar_status: { day3: 'pending', day10: 'pending', day30: 'pending' } },
  site_config: {},
  clash_requests: { status: 'pending' },
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

/**
 * The student page's prior-reflections check (filters: { student_id } when the student typed
 * their 800#, else { initial }): only what the page uses -- insight level, trait and date --
 * never names, answers or staff notes. No staff password needed (the student's own device).
 */
function studentHistory_(filters) {
  const f = filters.student_id ? { student_id: String(filters.student_id).replace(/\D/g, '') } : { initial: filters.initial };
  return genericQuery_('reflections', f).map(function (r) {
    const p3 = r.part3_json && typeof r.part3_json === 'object' ? r.part3_json : {};
    return { insight_level: r.insight_level || '', created_at: r.created_at, part3_json: { trait: p3.trait || '' } };
  });
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
      if (colIdx === -1) return false; // an unknown column matches nothing (it used to match everything)
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

/**
 * payload: { sessionId }. Lets an admin manually mark a submission as High insight even when
 * the AI scored it lower -- the AI's read is a starting point, not the final word, and Matt
 * may want to flag something the AI missed or under-scored. Also sets clash_flagged so the
 * row surfaces on the admin Clash Flags page exactly like an AI-scored High submission would,
 * making it eligible for a Clash of Classes point award. Gated admin-only -- see
 * enforceAdminGate_ in 05_AdminAuth.gs.
 */
function overrideInsightLevelHigh_(payload) {
  const sessionId = payload && payload.sessionId;
  if (!sessionId) throw new Error('sessionId is required.');
  genericUpdate_(TABS.REFLECTIONS, { session_id: sessionId }, { insight_level: 'High', clash_flagged: true });
  return { ok: true };
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
