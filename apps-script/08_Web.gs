/**
 * Trail Journal -- Apps Script backend
 * File 8: the pages themselves, served by this project (the GitHub site becomes thin frames).
 *
 * Pages (?page=): journal (default), respond (Your Side of the Trail), admin, panel (Trailback
 * Panel), followup (panel follow-up scheduler). Each page talks to tjCall() with
 * google.script.run instead of the public APP_TOKEN.
 *
 * Two deployments of this one project:
 *   Student link (Who has access: Anyone)  -- trailjournal.org and /respond/ show it in a frame.
 *   Staff link (Anyone within EASD)        -- /admin/ and /panel/ go straight to it. Staff are
 *     signed in with Google, so tjCall trusts their school account (STAFF_DOMAIN, optionally
 *     limited to STAFF_EMAILS) instead of the shared ADMIN_PASSWORD. The password still works
 *     as a fallback, and the per-adult check-in gate passwords are unchanged.
 * The old token route (doPost) stays on until the switch; LEGACY_API = OFF then limits it to the
 * exit auto-save beacon a closing journal sends.
 *
 * Settings tab (Key | Value | Notes): not readable through the API (only site_config is).
 */

const WEB_SETTINGS_TAB_ = 'Settings';
const WEB_DEFAULTS_ = [
  ['SITE_URL', 'https://trailjournal.org', 'The address students and staff use (links between pages go here).'],
  ['STAFF_DOMAIN', 'easdpa.org', 'Google accounts at this domain count as staff on the staff link.'],
  ['STAFF_EMAILS', '', 'Optional: only these staff emails (comma separated) get staff access. Blank = anyone at STAFF_DOMAIN.'],
  ['STUDENT_EXEC_URL', '', 'The student deployment\'s /exec link (Deploy › Manage deployments). The journal\'s exit auto-save goes there. Blank = this deployment.'],
  ['LEGACY_API', 'ON', 'ON while the old GitHub pages are still in use. Set OFF after the switch: the old token route then only takes the exit auto-save.'],
  ['PANEL_APPROVERS', '', 'Trailback Panel: emails (comma separated) allowed to approve a plan. Blank = anyone signed in on the staff link.'],
  ['TRAILBACK_CLASH_POINTS', '1', 'Trailback Panel: Clash of Classes points for a completed plan (after staff approve the request).'],
];
const PAGES_ = { journal: 'Journal', respond: 'Respond', admin: 'Admin', panel: 'Panel', followup: 'Followup' };

function webSetting_(key) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(WEB_SETTINGS_TAB_);
  if (sh) {
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) if (String(rows[i][0]).trim() === key) return String(rows[i][1] == null ? '' : rows[i][1]).trim();
  }
  const d = WEB_DEFAULTS_.filter(function (x) { return x[0] === key; })[0];
  return d ? d[1] : '';
}

function ensureWebSettings_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(WEB_SETTINGS_TAB_);
  if (!sh) {
    sh = ss.insertSheet(WEB_SETTINGS_TAB_);
    sh.getRange(1, 1, 1, 3).setValues([['Key', 'Value', 'Notes']]).setFontWeight('bold').setBackground('#2c4a35').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  const have = {};
  sh.getDataRange().getValues().slice(1).forEach(function (r) { have[String(r[0]).trim()] = true; });
  const missing = WEB_DEFAULTS_.filter(function (d) { return !have[d[0]]; });
  if (missing.length) sh.getRange(sh.getLastRow() + 1, 1, missing.length, 3).setValues(missing);
}

// The signed-in school staff account, or '' (students, signed-out visitors).
function staffEmail_() {
  let email = '';
  try { email = String(Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) { return ''; }
  const domain = webSetting_('STAFF_DOMAIN').toLowerCase();
  if (!email || !domain || email.slice(-(domain.length + 1)) !== '@' + domain) return '';
  const only = webSetting_('STAFF_EMAILS').toLowerCase().split(',').map(function (s) { return s.trim(); }).filter(String);
  return !only.length || only.indexOf(email) > -1 ? email : '';
}

// ── PAGES ───────────────────────────────────────────────────────────
function serveTrailPage_(e) {
  const p = (e && e.parameter) || {};
  const page = PAGES_[String(p.page || 'journal').toLowerCase()] ? String(p.page || 'journal').toLowerCase() : 'journal';
  const t = HtmlService.createTemplateFromFile(PAGES_[page]);
  let token = '';
  try { token = getProp_('APP_TOKEN'); } catch (err) { /* not set: exit auto-save just won't send */ }
  t.boot = JSON.stringify({
    page: page, staff: staffEmail_(), siteUrl: webSetting_('SITE_URL').replace(/\/+$/, ''),
    beaconUrl: webSetting_('STUDENT_EXEC_URL') || ScriptApp.getService().getUrl(), beaconToken: token,
    params: p,
  });
  return t.evaluate()
    .setTitle({ journal: 'Trail Journal — Camp Mountaineer', respond: 'Your Side of the Trail — Camp Mountaineer', admin: 'Trail Journal Admin',
      panel: 'Trailback Panel · Camp Mountaineer', followup: 'Trailback Panel Follow-Up' }[page])
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // trailjournal.org shows it in a frame
}

function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }

/** Every page's backend call (replaces fetch to the /exec with the public token). */
function tjCall(bodyJson) {
  try {
    const body = JSON.parse(bodyJson || '{}');
    delete body._staffEmail;
    const staff = staffEmail_();
    if (staff) body._staffEmail = staff;
    return JSON.stringify({ ok: true, data: dispatch_(body) });
  } catch (err) {
    return JSON.stringify({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

// ── SHEET MENU + SETUP ──────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Trail Journal')
    .addItem('Run setup (Settings tab + new columns)', 'setupTrailWeb')
    .addItem('Scramble gate passwords', 'scrambleGatePasswords')
    .addSeparator()
    .addItem('Turn on the record safety net', 'setupRecordSafetyNet') // 10_Records.gs
    .addItem('Email all journal records once…', 'emailAllRecordsOnce')
    .addItem('Send missed journal records one by one…', 'sendMissedRecords')
    .addToUi();
}

function setupTrailWeb() {
  ensureWebSettings_();
  addMissingColumns_(TABS.REFLECTIONS, ['student_id']);
  addMissingColumns_(TABS.RESPONDER_REFLECTIONS, ['student_name', 'student_id', 'share_name_with_other']);
  [TABS.PANEL_SESSIONS, TABS.FOLLOWUP_RECORDS, TABS.CLASH_REQUESTS].forEach(ensureSchemaTab_); // 11_Trailback.gs
  ensureNoSchoolTab_();
  safeAlert_('Trail Journal is set up: Settings tab (with PANEL_APPROVERS and TRAILBACK_CLASH_POINTS), the 800# columns, the new panel_sessions and followup_records columns, the clash_requests tab, and a No School Dates tab. Add the school\'s days off to No School Dates.');
}
