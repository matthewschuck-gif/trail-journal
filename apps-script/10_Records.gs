/**
 * Trail Journal -- Apps Script backend
 * File 10: every journal reaches the office and the Sheet, finished or not.
 *
 * WHY: the page showed "done" first and only sent the office email (and wrote the full report,
 * goal and Clash flag) after the AI summary came back. Start New Journal / Close Out reload the
 * page, and a Chromebook can sleep, so a long full-tier journal often lost all of that; short
 * detention journals usually made it. The server now does it itself:
 *   1. The AI summary call (02_AIProxy.gs) finishes the submission on the server: full report,
 *      goal and Clash flag on the row, and the office email -- even if the page is gone.
 *   2. The page's own office email is skipped when the server already sent that record.
 *   3. A safety net (every 10 minutes) sends anything still unsent: a submitted journal whose AI
 *      never ran gets the AI summary first; a partial journal is sent as INCOMPLETE.
 *   4. Menu › Send missed journal records… fills the Sheet and emails the records the office
 *      never got, back to any date.
 * The office_emailed column on reflections records what went out: "<time> submitted" or
 * "<time> incomplete", plus "resent" for records sent from the menu.
 */

const RECORD_SENT_COL_ = 'office_emailed';
const RECORD_WAIT_MIN_ = 15;    // the safety net leaves a row alone this long (the page may still be finishing)
const RECORD_AI_PER_RUN_ = 3;   // AI summaries are slow; at most this many per safety-net run
const RECORD_RUN_MS_ = 4.5 * 60 * 1000; // stop well before Apps Script's 6-minute limit

function parseJsonField_(v) {
  if (v && typeof v === 'object') return v;
  try { return v ? JSON.parse(String(v)) : {}; } catch (e) { return {}; }
}

// The reflections tab, with any SCHEMA column it's missing (office_emailed, last_saved_at,
// student_id on an older Sheet) added at the end in SCHEMA order -- the generic insert/update
// in 03_Router.gs write by SCHEMA position, so the Sheet has to match it.
function recordsTable_() {
  const sheet = getSheet_(TABS.REFLECTIONS);
  let headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0].map(function (h) { return String(h).trim(); });
  SCHEMA[TABS.REFLECTIONS].forEach(function (h) {
    if (headers.indexOf(h) !== -1) return;
    sheet.getRange(1, headers.length + 1).setValue(h);
    headers = headers.concat([h]);
  });
  const values = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues() : [];
  const col = function (name) { return headers.indexOf(name); };
  return {
    sheet: sheet, headers: headers, values: values, col: col,
    row: function (i) { const o = {}; headers.forEach(function (h, j) { o[h] = values[i][j]; }); return o; },
    set: function (i, name, value) { const c = col(name); if (c < 0) return; values[i][c] = value; sheet.getRange(i + 2, c + 1).setValue(value); },
  };
}

// ── EVERY SAVE FROM THE PAGE (router action upsertReflection) ───────────
// Progress at each screen ("in_progress"), Ran Out of Time / closing the tab ("incomplete") and
// the final submit ("submitted") all land on the same row, found by session_id, so a journal is
// on the Sheet from its first screen. A submitted journal is never turned back into a partial.
function upsertReflection_(row) {
  row = Object.assign({}, row || {});
  if (!row.session_id) throw new Error('upsertReflection needs a session_id');
  row.last_saved_at = nowIso_();
  const t = recordsTable_(), sc = t.col('session_id');
  let i = -1;
  for (let k = t.values.length - 1; k >= 0; k--) if (String(t.values[k][sc]) === String(row.session_id)) { i = k; break; }
  if (i < 0) { genericInsert_('reflections', row); return { inserted: true }; }
  if (String(t.row(i).status) === 'submitted' && row.status !== 'submitted') return { kept: 'submitted' };
  delete row.id; delete row.created_at; delete row[RECORD_SENT_COL_];
  genericUpdate_('reflections', { session_id: row.session_id }, row);
  return { updated: true };
}

const RECORD_IDLE_MIN_ = 60; // an in-progress journal untouched this long is treated as left unfinished

function lastSavedMs_(r) {
  const v = r.last_saved_at;
  const t = v instanceof Date ? v.getTime() : Date.parse(String(v || ''));
  return isNaN(t) ? rowCreatedMs_(r) : t;
}

function rowCreatedMs_(r) {
  const v = r.created_at;
  const t = v instanceof Date ? v.getTime() : Date.parse(String(v || ''));
  return isNaN(t) ? 0 : t;
}

// The page's S object, rebuilt from a saved row (what buildRecordText_ in 09_RecordText.gs reads).
function recordState_(r) {
  const cm = String(r.cm_reason || '');
  return {
    name: String(r.student_name || ''), initial: String(r.initial || ''), grade: String(r.grade || ''),
    squad: String(r.squad || ''), team: String(r.team || ''), location: String(r.location || ''), peak: String(r.peak || ''),
    cmReason: cm || null, cmInitiatedBy: r.cm_initiated_by || null, cmWhy: r.cm_why || null,
    cmEffortAgreed: r.cm_effort_agreed === true || String(r.cm_effort_agreed).toUpperCase() === 'TRUE',
    cmChecklistAgreed: r.cm_checklist_confirmed === true || String(r.cm_checklist_confirmed).toUpperCase() === 'TRUE',
    tier: (cm && CM_REASON_TIER_[cm]) || LOCATION_TIER_[String(r.location || '')] || 'full',
  };
}

function recordSentFor_(r, status) { return String(r[RECORD_SENT_COL_] || '').indexOf(status) !== -1; }

// Before tracking began, the page wrote reflection_text in the same step as each email, so a
// row with no report text never got its email. A finished journal still holding only the
// partial report (saved by Ran Out of Time, no insight level in it) never got its final email.
function recordLooksMissed_(r, status) {
  const text = String(r.reflection_text || '').trim();
  if (!text) return true;
  return status === 'submitted' && /Insight Level: —/.test(text);
}

// Writes what's missing on the row (full report, goal, Clash flag), sends the office email and
// marks the row. result = the AI summary (or {} when there is none).
function sendRecordForRow_(t, i, result, opts) {
  opts = opts || {};
  const r = t.row(i);
  const S = recordState_(r);
  const p1 = parseJsonField_(r.part1_json), p2 = parseJsonField_(r.part2_json), p3 = parseJsonField_(r.part3_json);
  result = result && !result._parseError ? result : {};
  const status = String(r.status || '') === 'submitted' ? 'submitted' : 'incomplete';
  const text = buildRecordText_(S, result, p1, p2, p3);
  // A finished journal's report replaces the partial one saved earlier; never blank out anything.
  if (!String(r.reflection_text || '').trim() || (status === 'submitted' && result.insightLevel)) t.set(i, 'reflection_text', text);
  if (result.studentGoal && !String(r.goal || '').trim()) t.set(i, 'goal', result.studentGoal);
  if (result.insightLevel === 'High' && String(r.clash_flagged || '') === '') t.set(i, 'clash_flagged', true);
  const created = rowCreatedMs_(r);
  sendOfficeRecordEmail_({
    studentLabel: S.name || S.initial || 'Student', grade: S.grade, location: S.location, tier: S.tier, squad: S.squad,
    incomplete: status !== 'submitted', summaryText: text, _fromServer: true,
    subjectNote: opts.resent ? 'resent' : '', submittedAt: created ? Utilities.formatDate(new Date(created), 'America/New_York', 'EEE M/d/yyyy h:mm a') : '',
  });
  t.set(i, RECORD_SENT_COL_, new Date().toISOString() + ' ' + status + (opts.resent ? ' resent' : ''));
  return status;
}

// Called by callAiProxy_ when a 'summary' finishes: the page may already be closed.
function finishSubmission_(sessionId, parsed) {
  const t = recordsTable_();
  const sc = t.col('session_id');
  for (let i = t.values.length - 1; i >= 0; i--) {
    if (String(t.values[i][sc]) !== String(sessionId)) continue;
    const r = t.row(i);
    if (String(r.status || '') !== 'submitted' || recordSentFor_(r, 'submitted')) return false;
    sendRecordForRow_(t, i, parsed);
    return true;
  }
  return false;
}

// The page's own office email (sendOfficeRecordEmail_ via the router) for the row it describes:
// the newest row in the last day for this student and location.
function matchRecordRow_(t, payload) {
  const label = String(payload.studentLabel || '').trim().toLowerCase();
  if (!label) return -1;
  const since = Date.now() - 24 * 3600 * 1000;
  for (let i = t.values.length - 1; i >= 0; i--) {
    const r = t.row(i);
    if (rowCreatedMs_(r) && rowCreatedMs_(r) < since) continue;
    const who = [String(r.student_name || '').trim().toLowerCase(), String(r.initial || '').trim().toLowerCase()];
    if (who.indexOf(label) === -1) continue;
    if (payload.location && String(r.location || '') !== String(payload.location)) continue;
    return i;
  }
  return -1;
}

// true = the page's email would repeat one the server already sent (a finished journal).
function recordAlreadySent_(payload) {
  if (payload._fromServer || payload.incomplete) return false;
  try {
    const t = recordsTable_(), i = matchRecordRow_(t, payload);
    return i > -1 && recordSentFor_(t.row(i), 'submitted');
  } catch (e) { return false; }
}

// After the page's own email goes out, mark the row so the safety net doesn't send it again.
function markRecordSentFromPage_(payload) {
  if (payload._fromServer) return;
  try {
    const t = recordsTable_(), i = matchRecordRow_(t, payload);
    if (i < 0) return;
    const r = t.row(i), status = payload.incomplete ? 'incomplete' : 'submitted';
    if (status === 'incomplete' && recordSentFor_(r, 'submitted')) return;
    t.set(i, RECORD_SENT_COL_, new Date().toISOString() + ' ' + status);
  } catch (e) { /* never block the email */ }
}

// The AI summary payload the page sends, rebuilt from a saved row.
function summaryPayloadFor_(r) {
  const S = recordState_(r);
  return { initial: S.initial, name: S.name, part1: parseJsonField_(r.part1_json), part2: parseJsonField_(r.part2_json), part3: parseJsonField_(r.part3_json),
    location: S.location, peak: S.peak, tier: S.tier, cmReason: S.cmReason, cmInitiatedBy: S.cmInitiatedBy, cmWhy: S.cmWhy,
    cmEffortAgreed: S.cmReason === CM_REASON_REQUIRES_EFFORT_ ? S.cmEffortAgreed : null,
    cmChecklistAgreed: S.cmReason === CM_REASON_REQUIRES_EFFORT_ ? S.cmChecklistAgreed : null,
    sessionId: String(r.session_id || ''), priorCount: Number(r.prior_count) || 0, priorReflections: [] };
}

// One pass over the rows. opts: { sinceMs, minAgeMs, onlyMissing, resent, aiLimit }.
// Returns { sent, ai, left } -- left > 0 means the time ran out; run again.
function sendUnsentRecords_(opts) {
  const start = Date.now(), t = recordsTable_();
  let sent = 0, ai = 0, left = 0;
  for (let i = 0; i < t.values.length; i++) {
    const r = t.row(i), created = rowCreatedMs_(r);
    if (!String(r.session_id || '').trim() || !created) continue;
    if (created < opts.sinceMs || Date.now() - created < opts.minAgeMs) continue;
    if (String(r.status || '') === 'in_progress') {
      // Still being written, or left open: wait until nobody has saved to it for an hour.
      if (Date.now() - lastSavedMs_(r) < RECORD_IDLE_MIN_ * 60 * 1000) continue;
      t.set(i, 'status', 'incomplete');
      if (!String(r.ran_out_of_time_by || '').trim()) t.set(i, 'ran_out_of_time_by', 'Auto (left unfinished)');
      r.status = 'incomplete';
    }
    const status = String(r.status || '') === 'submitted' ? 'submitted' : 'incomplete';
    const mark = String(r[RECORD_SENT_COL_] || '');
    if (opts.resent ? mark.indexOf('resent') !== -1 : recordSentFor_(r, status)) continue;
    if (opts.onlyMissing && !mark && !recordLooksMissed_(r, status)) {
      // The page wrote its report at the same moment it sent the email, before tracking began.
      t.set(i, RECORD_SENT_COL_, 'before tracking'); continue;
    }
    if (Date.now() - start > RECORD_RUN_MS_) { left++; continue; }
    let result = parseJsonField_(r.ai_staff_summary);
    if (status === 'submitted' && !result.insightLevel && ai < opts.aiLimit) {
      // sessionId '' so callAiProxy_ doesn't finish (and email) it too; this pass does that.
      try { result = callAiProxy_('summary', Object.assign(summaryPayloadFor_(r), { sessionId: '' })); ai++; } catch (e) { result = {}; }
      if (result && result.insightLevel && !result._parseError) updateReflectionAiFields_(String(r.session_id), result);
    }
    sendRecordForRow_(t, i, result, { resent: opts.resent });
    sent++;
    if (opts.resent) Utilities.sleep(1200); // spaced out so district mail doesn't throttle a burst
  }
  return { sent: sent, ai: ai, left: left };
}

// ── THE SAFETY NET (time trigger) ───────────────────────────────────────
function recordSafetyNet() {
  const since = Number(PropertiesService.getScriptProperties().getProperty('records_since')) || Date.now();
  sendUnsentRecords_({ sinceMs: since, minAgeMs: RECORD_WAIT_MIN_ * 60 * 1000, aiLimit: RECORD_AI_PER_RUN_ });
}

// Run once (menu or the editor's Run button): adds the column, starts tracking from now, and
// checks every 10 minutes for journals that never reached the office.
function setupRecordSafetyNet() {
  recordsTable_();
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('records_since')) props.setProperty('records_since', String(Date.now()));
  ScriptApp.getProjectTriggers().forEach(function (tr) { if (tr.getHandlerFunction() === 'recordSafetyNet') ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger('recordSafetyNet').timeBased().everyMinutes(10).create();
  recordAlert_('Record safety net is on. Every 10 minutes it emails (to the Notifications recipients) and fills in any journal, finished or partial, that never reached the office.');
}

// ── BACK-FILL (menu) ────────────────────────────────────────────────────
// Asks for a start date; "missing" (default) = only journals the office never got; ALL = every
// journal since that date again. Fills the Sheet the same way. Runs ~4.5 minutes, then says to
// run it again if more are left.
function sendMissedRecords() {
  const ui = SpreadsheetApp.getUi();
  const a = ui.prompt('Send missed journal records', 'Start date (YYYY-MM-DD). Records go to the office recipients on the Notifications page.', ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  const since = Date.parse(String(a.getResponseText() || '2026-08-01').trim() + 'T00:00:00');
  if (isNaN(since)) { ui.alert('That date did not read as YYYY-MM-DD.'); return; }
  const b = ui.prompt('Which journals?', 'Leave blank = only the ones the office never got (recommended).\nType ALL = every journal since that date, even ones already emailed.', ui.ButtonSet.OK_CANCEL);
  if (b.getSelectedButton() !== ui.Button.OK) return;
  const all = /^all$/i.test(String(b.getResponseText() || '').trim());
  const r = sendUnsentRecords_({ sinceMs: since, minAgeMs: 0, onlyMissing: !all, resent: true, aiLimit: 5 });
  ui.alert(r.sent + ' record' + (r.sent === 1 ? '' : 's') + ' emailed and filled in on the Sheet' + (r.ai ? ' (' + r.ai + ' got their AI summary first)' : '') + '.' +
    (r.left ? '\n\n' + r.left + ' more to go: run Send missed journal records again with the same answers.' : '\n\nAll caught up.'));
}

function recordAlert_(msg) { try { SpreadsheetApp.getUi().alert(msg); } catch (e) { console.log(msg); } }

// ── ONE-TIME PUSH (menu) ────────────────────────────────────────────────
// One email to the office recipients: every journal since a date as an attached spreadsheet
// (CSV, opens in Excel or Sheets), plus each journal the office never got, its full report
// written out in the email. Fills the Sheet's report column for those first and marks them so
// the safety net and the menu above don't send them again. Uses the AI summaries already on
// the Sheet; it doesn't run new ones (that could take longer than Apps Script allows).
function emailAllRecordsOnce() {
  const ui = SpreadsheetApp.getUi();
  const a = ui.prompt('Email all journal records once', 'Start date (YYYY-MM-DD), e.g. the first day of school.', ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  const since = Date.parse(String(a.getResponseText() || '').trim() + 'T00:00:00');
  if (isNaN(since)) { ui.alert('That date did not read as YYYY-MM-DD.'); return; }
  const r = emailAllRecords_(since);
  ui.alert('Sent to ' + getOfficeNotifyEmail_() + ':\n' + r.total + ' journals in the attached spreadsheet, ' + r.missed + ' that were never emailed written out in full.');
}

function emailAllRecords_(sinceMs) {
  const t = recordsTable_();
  const rows = [], missed = [];
  for (let i = 0; i < t.values.length; i++) {
    let r = t.row(i);
    const created = rowCreatedMs_(r);
    if (!String(r.session_id || '').trim() || created < sinceMs) continue;
    const status = String(r.status || '') === 'submitted' ? 'submitted' : 'incomplete';
    const result = parseJsonField_(r.ai_staff_summary);
    const S = recordState_(r);
    const wasMissed = !String(r[RECORD_SENT_COL_] || '') && recordLooksMissed_(r, status);
    if (wasMissed || !String(r.reflection_text || '').trim()) {
      t.set(i, 'reflection_text', buildRecordText_(S, result._parseError ? {} : result, parseJsonField_(r.part1_json), parseJsonField_(r.part2_json), parseJsonField_(r.part3_json)));
      if (result.studentGoal && !String(r.goal || '').trim()) t.set(i, 'goal', result.studentGoal);
      r = t.row(i);
    }
    if (wasMissed) { missed.push({ r: r, S: S, status: status }); t.set(i, RECORD_SENT_COL_, new Date().toISOString() + ' ' + status + ' resent once'); }
    else if (!String(r[RECORD_SENT_COL_] || '')) t.set(i, RECORD_SENT_COL_, 'before tracking');
    rows.push({ r: r, S: S, status: status, missed: wasMissed });
  }
  const when = function (r) { const c = rowCreatedMs_(r); return c ? Utilities.formatDate(new Date(c), 'America/New_York', 'yyyy-MM-dd h:mm a') : ''; };
  const cols = [['Started', function (x) { return when(x.r); }], ['Student', function (x) { return x.S.name || x.S.initial; }], ['800#', function (x) { return x.r.student_id || ''; }],
    ['Grade', function (x) { return x.S.grade; }], ['Squad', function (x) { return x.S.squad; }], ['Team', function (x) { return x.S.team; }],
    ['Location', function (x) { return x.S.location; }], ['Peaks Area', function (x) { return x.S.peak; }], ['Journal Length', function (x) { return x.S.tier; }],
    ['Status', function (x) { return x.status; }], ['Insight Level', function (x) { return x.r.insight_level || ''; }], ['Goal', function (x) { return x.r.goal || ''; }],
    ['Was Never Emailed', function (x) { return x.missed ? 'YES' : ''; }], ['Full Report', function (x) { return x.r.reflection_text || ''; }]];
  const cell = function (v) { v = String(v == null ? '' : v); return /[",\n\r]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; };
  const csv = '﻿' + [cols.map(function (c) { return c[0]; }).join(',')].concat(rows.map(function (x) { return cols.map(function (c) { return cell(c[1](x)); }).join(','); })).join('\r\n');
  const theme = squadTheme_('');
  const list = missed.length
    ? missed.map(function (m) {
        return '<div style="margin:18px 0;padding-top:12px;border-top:2px solid ' + theme.accent + '"><b>' + escHtml_(m.S.name || m.S.initial) + '</b> · ' + escHtml_(when(m.r)) + ' · ' +
          escHtml_(m.S.location) + (m.status === 'submitted' ? '' : ' · <b>INCOMPLETE</b>') + formatSummaryHtml_(String(m.r.reflection_text || ''), theme) + '</div>';
      }).join('')
    : '<p>Every journal in this range had already reached you.</p>';
  const body = emailCallout_('📎 Everything, once', rows.length + ' journals since ' + Utilities.formatDate(new Date(sinceMs), 'America/New_York', 'MMM d, yyyy') +
      ' are in the attached spreadsheet (every column, including the full written report). Below: the ' + missed.length + ' that never reached the office, written out in full. They are filled in on the Trail Journal Sheet too.', null, theme) + list;
  sendViaMail_({ to: getOfficeNotifyEmail_(), subject: '[Trail Journal] All journal records + ' + missed.length + ' that were missed', html: emailShell_('📎', 'Trail Journal Records', 'One-time catch-up', body, theme),
    attachments: [Utilities.newBlob(csv, 'text/csv', 'Trail Journal records since ' + Utilities.formatDate(new Date(sinceMs), 'America/New_York', 'yyyy-MM-dd') + '.csv')] });
  return { total: rows.length, missed: missed.length };
}

function escHtml_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
