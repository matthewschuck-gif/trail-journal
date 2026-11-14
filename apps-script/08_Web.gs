/**
 * Trail Journal -- Apps Script backend
 * File 8: the pages themselves, served by this project (the GitHub site becomes thin frames).
 *
 * Pages (?page=): journal (default), respond (Your Side of the Trail), admin, panel (Trailback
 * Panel), followup (panel follow-up scheduler), signup (panelists take a seat on a case). Each page talks to tjCall() with
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
  ['PANEL_ADMINS', '', 'Trailback Panel leads (comma separated emails): they set panel details and send briefs on the sign-up page. Blank = only the account that owns this script.'],
  ['PANEL_QUESTION_SHEET_URL', 'https://docs.google.com/document/d/15XwkkH-JA5m5jgz4UsZ0A0Obs3SoGkrUhCz3sSCrkTU/edit', 'The panelist question sheet, linked in the panel invite and the case brief. Blank = no link.'],
  ['PANEL_START_TIME', '12:00', 'Trailback Panel: the usual start time (24-hour, e.g. 12:00). A lead can change it per case on the sign-up page.'],
  ['PANEL_MINUTES', '25', 'Trailback Panel: how long a panel runs (minutes), for the calendar invite.'],
  ['PANEL_ROOM', '', 'Trailback Panel: the usual room. A lead can change it per case.'],
  ['PANEL_SEATS', '6', 'Trailback Panel: staff seats per case on the sign-up page.'],
  ['TRAILBACK_AUTOMATION', 'ON', 'Trailback autopilot (17_Pipeline.gs): ON = suggest cases from repeat patterns and Trail Journals, send the family consent link and the student\'s Trail Journal link, schedule when the journal is done, and call for panel seats. OFF = all by hand.'],
  ['TRAILBACK_PATTERN_COUNT', '3', 'Autopilot: this many Master Log rows with the same Behavior Type ...'],
  ['TRAILBACK_PATTERN_DAYS', '30', '... within this many days suggests a Trailback case.'],
  ['TRAILBACK_PATTERN_SKIP', 'bully, threat, weapon, safety, harass, drug, vape, fight', 'Autopilot: Behavior Types with any of these words are never suggested (they go to an administrator).'],
  ['TRAILBACK_TJ_PRIOR', '2', 'Autopilot: a Trail Journal with this many earlier journals (or a Low insight one with any before it) suggests a case.'],
  ['STUDENT_EMAIL_PATTERN', '{id}@gomounts.org', 'The student\'s school email, built from the 800# ({id}), for the Trail Journal link and panel date.'],
  ['TRAILBACK_AUTO_SCHEDULE', 'ON', 'Autopilot: ON = a case whose Trail Journal is done goes on the next Wednesday with room. OFF = staff schedule.'],
  ['PANEL_MAX_CASES', '2', 'Autopilot: the most cases on one Wednesday panel.'],
  ['PANEL_MIN_SEATS', '3', 'Autopilot: fewer staff than this signed up 2 school days before = a second seat call.'],
  ['TRAILBACK_AUTO_CLOSE', 'ON', 'Trailback: ON = when the check-in adult marks every action Done, the case closes as completed by itself (Master Log Resolved + the Clash point). OFF = staff close it. Nothing ever reverts by itself.'],
  ['LEGACY_API', 'ON', 'ON while the old GitHub pages are still in use. Set OFF after the switch: the old token route then only takes the exit auto-save.'],
  ['PANEL_APPROVERS', '', 'Trailback Panel: emails (comma separated) allowed to approve a plan. Blank = anyone signed in on the staff link.'],
  ['TRAILBACK_CLASH_POINTS', '1', 'Trailback Panel: Clash of Classes points for a completed plan (after staff approve the request).'],
  ['TRAILBACK_EMAILS', 'OFF', 'Trailback Panel: ON sends the Tuesday brief, the overdue check-in email and the reverted-case email. OFF sends nothing (the menu previews still work).'],
  ['TRAILBACK_COORDINATOR_EMAIL', '', 'Trailback Panel: Kelly\'s email (comma separated for more). Gets overdue check-ins, reverted cases, and conflict-of-interest replies.'],
  ['TRAILBACK_CONFLICT_DEADLINE', 'Wednesday morning, before the panel', 'Trailback Panel: when panelists must tell Kelly about a conflict of interest (shown in the Tuesday brief).'],
  ['ACCOUNTABILITY_SHEET_ID', '1IVFnkQQvSvNSY_OZPMutPzkTqmJ1j5XYRoT9YvNy-v4', 'Trailback Panel: the Accountability & Intervention Log Sheet. Its Master Log Status changes when a case is held, resolved or reverted.'],
  ['TRAILBACK_REFERRAL_FORM_ID', '19qZF3mndYFktrkrwHZz-lr5IGSJUAoU46w99rUTShvc', 'Trailback Panel: the "Trailback Panel Staff Referral" Google Form. Connect it from the menu.'],
];
const PAGES_ = { journal: 'Journal', respond: 'Respond', admin: 'Admin', panel: 'Panel', followup: 'Followup', signup: 'Signup', checkin: 'Checkin', givget: 'GiveGet', consent: 'Consent' };

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
      panel: 'Trailback Panel · Camp Mountaineer', followup: 'Trailback Panel Follow-Up', signup: 'Trailback Panel Sign-Up', checkin: 'Trailback Check-In', givget: 'Give / Get · EMS', consent: 'Trailback Panel · Ephrata Middle School' }[page])
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
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Trailback') // 12_TrailbackCases.gs
      .addItem('Turn on the case briefs (3:00 the school day before) + overdue check', 'setupTrailbackTriggers')
      .addItem('Preview the next case briefs (nothing sent)', 'previewPanelBriefs')
      .addItem('Open the panel sign-up page', 'showPanelSignupLink')
      .addItem('Open the one-tap check-in page', 'showCheckinLink')
      .addItem('Open Give / Get for staff', 'showGiveGetLink')
      .addItem('Run the autopilot now (suggest, remind, schedule, seats)', 'menuRunPipeline_')
      .addItem('Preview the overdue check-in email (nothing sent)', 'previewOverdueEmail')
      .addItem('Add missing questions to the staff referral form', 'addReferralQuestions')
      .addItem('Connect the staff referral form', 'connectReferralForm')
      .addItem('Show the key for other tools', 'showTrailbackKey')
      .addSeparator()
      .addItem('Add a test case', 'addTrailbackTestCase')
      .addItem('Remove test cases', 'removeTrailbackTestCases'))
    .addToUi();
}

function setupTrailWeb() {
  ensureWebSettings_();
  addMissingColumns_(TABS.REFLECTIONS, ['student_id']);
  addMissingColumns_(TABS.RESPONDER_REFLECTIONS, ['student_name', 'student_id', 'share_name_with_other']);
  [TABS.PANEL_SESSIONS, TABS.FOLLOWUP_RECORDS, TABS.CLASH_REQUESTS, TABS.PANEL_CASES].forEach(ensureSchemaTab_); // 11_Trailback.gs
  ensurePanelQuestionsTab_(); // 14_RepairPlan.gs: the v2 question sheet, editable
  ensureGiveGetTabs_(); // 15_GiveGet.gs: the Action Ideas bank and the Give / Get board
  ensureNoSchoolTab_();
  ensurePanelMembersTab_(); // 12_TrailbackCases.gs
  safeAlert_('Trail Journal is set up: Settings tab (with PANEL_APPROVERS and TRAILBACK_CLASH_POINTS), the 800# columns, the new panel_sessions and followup_records columns, the clash_requests tab, and a No School Dates tab. Add the school\'s days off to No School Dates.');
}
