/**
 * Trail Journal -- Apps Script backend
 * File 11: the Trailback Panel's own rules.
 *
 *   - School-day follow-up dates: every panel case is checked in on at 3, 10 and 30 school
 *     days after the panel. Weekends and the dates on the "No School Dates" tab don't count.
 *     The Panel page and the follow-up scheduler both ask the server (followupDates), so they
 *     always agree.
 *   - AP approval: approvePanelPlan records who approved (the signed-in Google account on the
 *     staff link, never a value the page sends) and when. submitPanelSession refuses a session
 *     that hasn't been approved.
 *   - Initials: the AI helpers get first + last initial only, never the name.
 *
 * Settings tab rows used here (08_Web.gs adds them on Run setup):
 *   PANEL_APPROVERS        comma-separated emails allowed to approve. Blank = any signed-in staff.
 */

const NO_SCHOOL_TAB_ = 'No School Dates';
const CHECKIN_SCHOOL_DAYS_ = [3, 10, 30];
// Set only by approvePanelPlan_; the generic insert/update actions drop them (03_Router.gs).
const PANEL_PROTECTED_FIELDS_ = ['approved_by', 'approved_at'];

// ── NAMES ───────────────────────────────────────────────────────────
/** "Ann Lee" -> "A.L.", "Mary-Kate O'Neil" -> "M.O.", "Cher" -> "C." */
function initialsOf_(name) {
  const parts = String(name || '').trim().split(/\s+/).filter(String);
  if (!parts.length) return '';
  const first = parts[0].charAt(0).toUpperCase();
  const last = parts.length > 1 ? parts[parts.length - 1].charAt(0).toUpperCase() : '';
  return first + '.' + (last ? last + '.' : '');
}

// ── SCHOOL DAYS ─────────────────────────────────────────────────────
function ymdOf_(v) {
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return '';
    return v.getFullYear() + '-' + String(v.getMonth() + 1).padStart(2, '0') + '-' + String(v.getDate()).padStart(2, '0');
  }
  const s = String(v || '').trim();
  const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + m[2].padStart(2, '0') + '-' + m[3].padStart(2, '0');
  const us = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); // 12/24/2026, as typed into the tab
  if (us) return us[3] + '-' + us[1].padStart(2, '0') + '-' + us[2].padStart(2, '0');
  return '';
}

/** The dates on the No School Dates tab, as a { 'yyyy-MM-dd': true } set. */
function noSchoolDates_() {
  const set = {};
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NO_SCHOOL_TAB_);
  if (!sh) return set;
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    const d = ymdOf_(r[0]);
    if (d) set[d] = true;
  });
  return set;
}

/**
 * The date n school days after fromStr (yyyy-MM-dd). The panel day itself is day 0, so a
 * Wednesday panel's day 3 is the following Monday when there's no day off in between.
 * Weekends and noSchool dates are skipped.
 */
function schoolDaysAfter_(fromStr, n, noSchool) {
  const ymd = ymdOf_(fromStr);
  if (!ymd) throw new Error('A date (yyyy-MM-dd) is needed to count school days.');
  const p = ymd.split('-').map(Number);
  const d = new Date(Date.UTC(p[0], p[1] - 1, p[2], 12));
  const off = noSchool || {};
  let count = 0;
  while (count < n) {
    d.setUTCDate(d.getUTCDate() + 1);
    const dow = d.getUTCDay();
    const key = d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' + String(d.getUTCDate()).padStart(2, '0');
    if (dow === 0 || dow === 6 || off[key]) continue;
    count++;
  }
  return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' + String(d.getUTCDate()).padStart(2, '0');
}

/** payload: { from }. -> { checkin_3, checkin_10, checkin_30 } (the panel's follow-up dates). */
function followupSchedule_(payload) {
  const from = ymdOf_(payload && payload.from) || ymdOf_(new Date());
  const off = noSchoolDates_();
  const out = { from: from };
  CHECKIN_SCHOOL_DAYS_.forEach(function (n) { out['checkin_' + n] = schoolDaysAfter_(from, n, off); });
  return out;
}

/**
 * Old panel rows saved followup_dates as { two_days, one_week, two_weeks, one_month }; new
 * rows save { checkin_3, checkin_10, checkin_30 }. Either way -> [{ key, label, date }].
 */
function panelCheckins_(dates) {
  const d = dates || {};
  if (d.checkin_3 || d.checkin_10 || d.checkin_30) {
    return CHECKIN_SCHOOL_DAYS_.map(function (n) { return { key: 'checkin_' + n, label: n + ' school days', date: d['checkin_' + n] || '' }; });
  }
  return [['two_days', '2 days'], ['one_week', '1 week'], ['two_weeks', '2 weeks'], ['one_month', '1 month']]
    .map(function (x) { return { key: x[0], label: x[1], date: d[x[0]] || '' }; });
}

function ensureNoSchoolTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(NO_SCHOOL_TAB_);
  if (sh) return sh;
  sh = ss.insertSheet(NO_SCHOOL_TAB_);
  sh.getRange(1, 1, 1, 2).setValues([['Date', 'Note']]).setFontWeight('bold').setBackground('#2c4a35').setFontColor('#ffffff');
  sh.getRange(2, 3).setValue('One date per row (any date format). Weekends never count, so only list weekdays off: holidays, in-service days, snow days.');
  sh.setFrozenRows(1);
  return sh;
}

// ── SHEET COLUMNS ───────────────────────────────────────────────────
/** Makes sure a tab exists and has every SCHEMA column (missing ones go on the end, in order). */
function ensureSchemaTab_(tab) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(tab);
  const want = SCHEMA[tab];
  if (!sh) {
    sh = ss.insertSheet(tab);
    sh.getRange(1, 1, 1, want.length).setValues([want]).setFontWeight('bold').setBackground('#2c4a35').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    return sh;
  }
  const have = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(function (h) { return String(h).trim(); });
  const missing = want.filter(function (h) { return have.indexOf(h) === -1; });
  if (missing.length) {
    const start = have.filter(String).length + 1;
    sh.getRange(1, start, 1, missing.length).setValues([missing]).setFontWeight('bold').setBackground('#2c4a35').setFontColor('#ffffff');
  }
  return sh;
}

// ── APPROVAL ────────────────────────────────────────────────────────
function canApprovePanel_(email) {
  const e = String(email || '').toLowerCase();
  if (!e) return false;
  const list = String(webSetting_('PANEL_APPROVERS') || '').toLowerCase().split(',').map(function (s) { return s.trim(); }).filter(String);
  return !list.length || list.indexOf(e) > -1;
}

function stripPanelApproval_(obj) {
  if (!obj || typeof obj !== 'object') return obj;
  const out = Object.assign({}, obj);
  PANEL_PROTECTED_FIELDS_.forEach(function (k) { delete out[k]; });
  // Approval statuses come only from approvePanelPlan_ / submitPanelSession_.
  if (out.status === 'heard_approved' || out.status === 'repair_in_progress') delete out.status;
  return out;
}

function stripProtected_(obj) {
  const out = Object.assign({}, obj || {});
  PANEL_PROTECTED_FIELDS_.forEach(function (k) { delete out[k]; });
  delete out.id; delete out.created_at; delete out.status;
  return out;
}

function panelRow_(sessionId) {
  return genericQuery_(TABS.PANEL_SESSIONS, { session_id: sessionId }, 1)[0] || null;
}

/**
 * body.payload: { record } (the session so far, with session_id). Signed-in staff on the
 * staff link only. Saves the session as heard_approved with approved_by = the Google
 * account. Approving twice keeps the first approval.
 */
function approvePanelPlan_(body) {
  const email = body._staffEmail;
  if (!email) throw new Error('Approving a plan needs a school Google sign-in. Open the panel from the staff link (trailjournal.org/panel).');
  if (!canApprovePanel_(email)) throw new Error(email + ' is not on the PANEL_APPROVERS list (Settings tab).');
  const record = (body.payload && body.payload.record) || {};
  const sid = String(record.session_id || '');
  if (!sid) throw new Error('session_id is required.');
  ensureSchemaTab_(TABS.PANEL_SESSIONS);

  const existing = panelRow_(sid);
  if (existing && existing.approved_by) return { approved_by: existing.approved_by, approved_at: existing.approved_at, status: existing.status };

  const approval = { approved_by: email, approved_at: nowIso_(), status: 'heard_approved' };
  if (existing) genericUpdate_(TABS.PANEL_SESSIONS, { session_id: sid }, Object.assign(stripProtected_(record), approval));
  else genericInsert_(TABS.PANEL_SESSIONS, Object.assign(stripProtected_(record), { session_id: sid }, approval));
  return approval;
}

/**
 * body.payload: { record }. Saves the finished session. Refuses unless approvePanelPlan_ ran
 * for this session_id. Status moves to repair_in_progress; the approval is never changed here.
 */
function submitPanelSession_(body) {
  const record = (body.payload && body.payload.record) || {};
  const sid = String(record.session_id || '');
  if (!sid) throw new Error('session_id is required.');
  const existing = panelRow_(sid);
  if (!existing || !existing.approved_by) throw new Error('This plan has not been approved yet. The AP approves it before the record is submitted.');
  genericUpdate_(TABS.PANEL_SESSIONS, { session_id: sid }, Object.assign(stripProtected_(record), { status: 'repair_in_progress' }));
  return Object.assign({}, panelRow_(sid));
}
