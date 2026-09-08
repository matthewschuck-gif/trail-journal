/**
 * Trail Journal -- Apps Script backend
 * File 1 of N: Config + one-time spreadsheet provisioning
 *
 * SETUP STEPS (do these once, in order):
 * 1. Create a new Google Sheet in your school Workspace Drive (any name, e.g. "Trail Journal Data").
 * 2. Open it -> Extensions -> Apps Script. This creates a bound Apps Script project.
 * 3. Paste each 0N_*.gs file from this package into the Apps Script editor as its own file
 *    (File -> New -> Script file, name it to match, e.g. "01_Config").
 * 4. In the Apps Script editor: Project Settings -> Script Properties -> add:
 *      ANTHROPIC_API_KEY   = sk-ant-...          (Matt adds this directly, never in chat)
 *      APP_TOKEN            = <any random string you make up> (shared token, see 03_Router.gs)
 *      ADMIN_PASSWORD       = <a password for staff> (checked server-side, see 05_AdminAuth.gs)
 *      OFFICE_NOTIFY_EMAIL  = <office/record-keeping inbox> (OPTIONAL -- every submission,
 *                             every tier/pathway, automatically emails a summary copy here
 *                             for the record. Defaults to emsoffice@easdpa.org if unset --
 *                             see getOfficeNotifyEmail_() in 04_FollowupEmails.gs. This is
 *                             separate from, and in addition to, the staff-scheduled
 *                             follow-up email.)
 *    (No Resend key or FROM_EMAIL needed -- 04_FollowupEmails.gs sends via MailApp, Apps
 *    Script's own mail service, using whichever Google account deploys this project. See
 *    that file's header for why this replaced Resend.)
 * 5. Run `setupSpreadsheet` once from the editor (select it in the function dropdown, click Run).
 *    Approve the permission prompts -- this first run is also where you'll be asked to
 *    authorize the "send email as you" permission MailApp needs. That consent screen is
 *    shown ONLY to you, the developer, this one time; students and staff using the actual
 *    site are never shown any Google sign-in or consent prompt -- the deployed Web App always
 *    executes as your account automatically (see step 6), so the follow-up emails and
 *    calendar invites go out with no login step for whoever's using the page.
 * 6. Deploy ONCE as a Web App: Execute as "Me", Who has access "Anyone".
 *    (An earlier version of this used two deployments -- one domain-restricted for staff --
 *    but Apps Script's domain-restricted deployments don't return CORS headers to
 *    cross-origin fetch() calls, so that path never actually worked from trailjournal.org.
 *    One public deployment + a server-side password check in 05_AdminAuth.gs replaces it.)
 *    Copy the resulting /exec URL -- it goes in ALL FIVE frontend files.
 */

// --- SHEET TAB NAMES (mirrors the original Supabase table names 1:1) ---
const TABS = {
  REFLECTIONS: 'reflections',
  RESPONDER_REFLECTIONS: 'responder_reflections',
  PANEL_SESSIONS: 'panel_sessions',
  FOLLOWUP_RECORDS: 'followup_records',
  INCIDENT_REPORTS: 'incident_reports',
  SITE_CONFIG: 'site_config',
  STAFF_USERS: 'staff_users',
};

// Exact column order per tab -- matches the original Postgres schema field-for-field.
// json/jsonb columns are stored as JSON-stringified text in a single cell.
const SCHEMA = {
  [TABS.REFLECTIONS]: [
    'id', 'created_at', 'session_id', 'initial',
    'part1_json', 'part2_json', 'part3_json',
    'insight_level', 'ai_staff_summary', 'ai_reading_list', 'has_flags',
    'status', 'goal', 'prior_count', 'conference_brief',
    'panel_completed', 'panel_session_id', 'panel_recommendation', 'panel_lrg_trait',
    'panel_checkin_person', 'panel_checkin_date', 'panel_completed_at',
    // Added when the journal switched from anonymous-initial to full-name collection
    // (data is now used for intervention/support, not anonymized) -- appended at the
    // END of the column list, not inserted earlier, so existing rows/columns in an
    // already-provisioned Sheet are not shifted. See migrateAddNameGradeColumns_() below
    // if you already ran setupSpreadsheet() before this change.
    'student_name', 'grade',
    // Added for the tiered-journal-by-consequence-severity feature: 'location' is where
    // the student is completing the journal (Lunch/Morning/Afterschool Detention /
    // Office / ISS / Camp Mountaineer / Other -- see LOCATION_TIER_ in index.html), which
    // determines whether they get the full 3-part journal, the 2-part version, or the
    // shortest office version. 'peak' is which Mountaineer Peaks universal expectation
    // (Be Present / Be Personable / Be Productive) the incident relates to most. Both
    // appended at the end -- see migrateAddLocationPeakColumns_() below.
    'location', 'peak',
    // Added for the Camp Mountaineer pre-page (only shown when location is 'Camp
    // Mountaineer'): 'cm_reason' is which of the 4 pathways brought the student in
    // (Behavior Reflection - Time in Office / Behavior Reflection - Alternative to ISS /
    // Scheduled Break or Reflection Check-In / Re-Entry), 'cm_initiated_by' is
    // Student-Initiated or Staff-Initiated, 'cm_why' is the free-text reason, and
    // 'cm_effort_agreed' is true only when the Alternative-to-ISS pathway's Effort
    // Agreement checklist (from the Camp Mountaineer Updates doc) was fully confirmed --
    // null for every other pathway, where it isn't required. 'cm_checklist_confirmed' is
    // true when the general Camp Mountaineer Student Checklist (same doc) was fully
    // confirmed -- applies to every pathway, unlike the Effort Agreement. See
    // CM_REASON_TIER_, CM_STUDENT_CHECKLIST_ITEMS_, CM_EFFORT_AGREEMENT_ITEMS_ and
    // screen-cm in index.html. Appended at the end -- see
    // migrateAddCampMountaineerColumns_() below.
    'cm_reason', 'cm_initiated_by', 'cm_why', 'cm_effort_agreed', 'cm_checklist_confirmed',
    // Which staff member's individual password unlocked each check-in gate (see
    // staff_users / checkStaffPassword_ in 05_AdminAuth.gs) -- the accountability trail
    // that replacing the shared password was for. gate1/gate2 are null for tiers that
    // skip that gate (office has neither; detention has only gate1). Appended at the end --
    // see migrateAddStaffGateColumns_() below.
    'gate1_unlocked_by', 'gate2_unlocked_by', 'final_unlocked_by',
    // Flattened, human-readable version of the entire reflection -- every Part 1/2/3
    // answer verbatim, same content as the "FULL REFLECTION" section of the automatic
    // office-record email (buildEmailBody() in index.html; see blockFullReflection there).
    // Written once, right after the AI summary comes back (see the reflection_text patch
    // in doSubmit()), so the raw part1_json/part2_json/part3_json blobs always have a
    // plain-text sibling column that's actually readable/searchable/exportable straight out
    // of the Sheet -- no JSON parsing needed to see what a student actually wrote. Appended
    // at the end, same safe pattern as every other migrate*_ helper below.
    'reflection_text',
    // Clash of Classes bookkeeping on the Trail Journal's own row -- separate from the
    // actual point award, which awardClashPoint_ (07_ClashOfClasses.gs) writes straight to
    // the real clashofclasses.org backend (a separate Google Sheet's "events" tab) once a
    // staff member approves it on the admin Clash Flags page. These two columns exist so
    // THIS record can say "yes, a point was already given for this row" -- without them,
    // the admin Clash Flags page's clash_flagged=eq.true filter silently matched nothing
    // was ever filtered out (genericQuery_ ignores a filter on a column that doesn't
    // exist), and its "Award Point" button's patch silently no-opped (genericUpdate_ skips
    // unknown columns) --
    // meaning the page listed every submission, not just flagged ones, and could never
    // actually record a manual award. See migrateAddClashFlagColumns_() below.
    'clash_flagged', 'clash_points_awarded',
    // The student's Clash of Classes squad (BLACK/GOLD/GREY/PURPLE), picked on the welcome
    // screen -- previously only ever held in the browser's in-memory S.squad, never saved
    // anywhere server-side. Needed now so an admin reviewing this row on the Clash Flags
    // page later (see awardClashPoint_ in 07_ClashOfClasses.gs) knows which squad to award
    // the point to -- without this column there was no way to recover that after the fact.
    'squad',
    // Which team referred this consequence (M/O/U/N/T/S/Building MTSS Team) -- collected
    // on the welcome screen for every location now (not just Removal from Lunch, where it
    // started -- see the paper form that came from). Top-level like squad/location/peak,
    // not nested in part1_json, since it's demographic info, not Part 1 content.
    'team',
    // Records which staff member unlocked the NEW entry gate (pw-gate-0 in index.html) --
    // added between the welcome/demographic screen and Part 1 so an adult must be present
    // from the very start, not just at the two existing mid-journal check-ins. Same
    // accountability-trail pattern as gate1_unlocked_by/gate2_unlocked_by above.
    'gate0_unlocked_by',
    // Which staff member authorized an early/incomplete save via the "Ran out of time"
    // button (see submitRanOutOfTime() in index.html) -- null for every normal completed
    // submission. Lets admin tell a genuinely-finished reflection apart from one a student
    // didn't get to finish in the period, and who approved cutting it short.
    'ran_out_of_time_by',
  ],
  [TABS.RESPONDER_REFLECTIONS]: [
    'id', 'created_at', 'session_id', 'linked_session_id', 'initial',
    'what_happened', 'how_affected', 'what_hardest',
    'emotions_json', 'intensity', 'what_needs_to_happen', 'ready_to_talk', 'status',
  ],
  [TABS.PANEL_SESSIONS]: [
    'id', 'session_id', 'created_at', 'panel_date', 'student_initial', 'referring_staff',
    'incident_summary', 'behavior_pattern', 'prior_interventions', 'trail_journal_id',
    'student_what_happened', 'student_who_affected', 'student_thought_since', 'student_what_needs',
    'panel_questions', 'impact_statements',
    'deliberation_harm', 'deliberation_student_needs', 'deliberation_community_needs',
    'recommendation_categories', 'recommendation_narrative', 'lrg_trait',
    'followup_person', 'followup_dates', 'followup_notes', 'status',
    'clash_squad', 'clash_point_awarded', 'ai_recommendation', 'ai_summary',
    // Added alongside the reflections student_name/grade change -- 'student_initial' now
    // holds the full name from the Panel's "Student Name" field (the column wasn't renamed
    // to avoid a disruptive Sheet header change). Appended at the end for the same reason.
    'grade',
  ],
  [TABS.FOLLOWUP_RECORDS]: [
    'id', 'created_at', 'student_initials', 'trigger_type', 'trigger_date', 'outcome_summary',
    'panel_session_id', 'attendee_emails',
    'checkin_2day_date', 'checkin_2week_date', 'checkin_2month_date', 'calendar_status',
    'submitted_by', 'squad', 'notes',
    // Added so single-event Trail Journal follow-ups (trigger_type = 'trail-journal',
    // scheduled from scheduleFuCheckin() in index.html) can actually be closed out --
    // previously there was no admin UI for these at all (the existing Panel Follow-Up
    // Tracker page only ever queried panel_sessions, a different table). See the
    // Trail Journal Follow-Ups section in admin/index.html (page-followup) for the UI
    // that reads/writes these. Appended at the end -- see
    // migrateAddFollowupCompletionColumns_() below.
    'completed', 'completed_at', 'completed_by', 'completion_notes',
  ],
  [TABS.INCIDENT_REPORTS]: [
    'id', 'created_at', 'student_name', 'grade', 'homeroom',
    'what_happened', 'who_was_involved', 'student_perspective',
    'status', 'assigned_to', 'staff_notes', 'follow_up_action', 'closed_at', 'closed_by',
  ],
  [TABS.SITE_CONFIG]: ['key', 'value', 'updated_at'],
  // Individual staff passwords for the mid-journal check-in gates and the final screen
  // unlock -- replaces the single shared ADMIN_PASSWORD for those three specific unlock
  // points so each gate records WHICH staff member approved it. The admin DASHBOARD login
  // still uses ADMIN_PASSWORD separately (05_AdminAuth.gs) -- kept intentionally distinct
  // since dashboard access is a more sensitive surface than "an adult supervised this
  // check-in." Passwords are plain text in this sheet, same trust model as ADMIN_PASSWORD
  // living in Script Properties -- protect this tab's sharing settings accordingly. See
  // migrateAddStaffUsersTab_() below for one-time setup and the starting name list.
  [TABS.STAFF_USERS]: ['name', 'password'],
};

/**
 * incident_reports contains real student names, grades, and homerooms -- this is the one
 * tab with actual PII (not anonymized like the other student-facing flows). Once created,
 * go to that sheet tab -> right-click -> Protect range, and restrict edit/view access to the
 * same staff who currently have admin-dashboard access. Do not share this spreadsheet broadly.
 */

/**
 * Run this ONCE, only if you already ran setupSpreadsheet() before student_name/grade (and
 * panel_sessions.grade) were added to the schemas above. Safely appends the new header
 * cells to the END of each tab's existing header row -- does not touch or reorder any
 * existing columns or data. Safe to run more than once (skips columns that already exist).
 */
function migrateAddNameGradeColumns_() {
  const results = [];
  results.push(addMissingColumns_(TABS.REFLECTIONS, ['student_name', 'grade']));
  results.push(addMissingColumns_(TABS.PANEL_SESSIONS, ['grade']));
  safeAlert_(results.join('\n'));
}

/**
 * Run this ONCE, only if you already ran setupSpreadsheet() before location/peak were
 * added to the reflections schema above. Same safe append-only behavior as
 * migrateAddNameGradeColumns_().
 */
function migrateAddLocationPeakColumns_() {
  const result = addMissingColumns_(TABS.REFLECTIONS, ['location', 'peak']);
  safeAlert_(result);
}

/**
 * Run this ONCE, only if you already ran setupSpreadsheet() before the Camp Mountaineer
 * pre-page columns were added to the reflections schema above. Same safe append-only
 * behavior as the other migrate*_ helpers.
 */
function migrateAddCampMountaineerColumns_() {
  const result = addMissingColumns_(TABS.REFLECTIONS, ['cm_reason', 'cm_initiated_by', 'cm_why', 'cm_effort_agreed', 'cm_checklist_confirmed']);
  safeAlert_(result);
}

/**
 * Run this ONCE, only if you already ran setupSpreadsheet() before the follow-up
 * completion-tracking columns were added to the followup_records schema above.
 */
function migrateAddFollowupCompletionColumns_() {
  const result = addMissingColumns_(TABS.FOLLOWUP_RECORDS, ['completed', 'completed_at', 'completed_by', 'completion_notes']);
  safeAlert_(result);
}

/**
 * Run this ONCE, only if you already ran setupSpreadsheet() before the gate-unlock
 * tracking columns were added to the reflections schema above.
 */
function migrateAddStaffGateColumns_() {
  const result = addMissingColumns_(TABS.REFLECTIONS, ['gate1_unlocked_by', 'gate2_unlocked_by', 'final_unlocked_by']);
  safeAlert_(result);
}

/**
 * Run this ONCE, only if you already ran setupSpreadsheet() before reflection_text was
 * added to the reflections schema above. Same safe append-only behavior as the other
 * migrate*_ helpers.
 */
function migrateAddReflectionTextColumn_() {
  const result = addMissingColumns_(TABS.REFLECTIONS, ['reflection_text']);
  safeAlert_(result);
}

/**
 * Run this ONCE, only if you already ran setupSpreadsheet() before clash_flagged/
 * clash_points_awarded were added to the reflections schema above. Same safe append-only
 * behavior as the other migrate*_ helpers -- fixes the Clash Flags admin page (see the
 * comment on these two columns in SCHEMA above for what was broken without them).
 */
function migrateAddClashFlagColumns_() {
  const result = addMissingColumns_(TABS.REFLECTIONS, ['clash_flagged', 'clash_points_awarded']);
  safeAlert_(result);
}

/**
 * Run this ONCE, only if you already ran setupSpreadsheet() before squad was added to the
 * reflections schema above. Same safe append-only behavior as the other migrate*_ helpers.
 */
function migrateAddSquadColumn_() {
  const result = addMissingColumns_(TABS.REFLECTIONS, ['squad']);
  safeAlert_(result);
}

/**
 * Run this ONCE, only if you already ran setupSpreadsheet() before team was added to the
 * reflections schema above. Same safe append-only behavior as the other migrate*_ helpers.
 */
function migrateAddTeamColumn_() {
  const result = addMissingColumns_(TABS.REFLECTIONS, ['team']);
  safeAlert_(result);
}

/**
 * Run this ONCE, only if you already ran setupSpreadsheet() before gate0_unlocked_by and
 * ran_out_of_time_by were added to the reflections schema above. Same safe append-only
 * behavior as the other migrate*_ helpers.
 */
function migrateAddGate0AndRanOutOfTimeColumns_() {
  const result = addMissingColumns_(TABS.REFLECTIONS, ['gate0_unlocked_by', 'ran_out_of_time_by']);
  safeAlert_(result);
}

/**
 * Run this ONCE to create the staff_users tab and seed it with your starting staff list.
 * Creates the tab if it doesn't exist yet, and only seeds names if the tab is empty (safe
 * to re-run -- it will never overwrite existing rows). Passwords are left BLANK on purpose:
 * go to the staff_users tab afterward and type a password into column B next to each name
 * yourself. "Guest" is meant to be shared with any adult not on the list -- give it a
 * password too and anyone can use it, still distinct from every named staff member's own.
 */
function migrateAddStaffUsersTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(TABS.STAFF_USERS);
  if (!sheet) sheet = ss.insertSheet(TABS.STAFF_USERS);
  const headers = SCHEMA[TABS.STAFF_USERS];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#2c4a35').setFontColor('#ffffff');

  if (sheet.getLastRow() < 2) {
    const names = ['Mr. Miller', 'Mrs. Rigg', 'Mr. Schuck', 'Guest', 'Mrs. Mincarelli',
      'Mrs. Mowbray', 'Mrs. Lugar', 'Dr. Montagna', 'Mrs. Wagner', 'Mrs. Judge', 'Mr. Kuhn'];
    sheet.getRange(2, 1, names.length, 1).setValues(names.map(function (n) { return [n]; }));
  }
  safeAlert_('staff_users tab ready. Now go type a password into column B next to each name before that person can unlock a gate -- passwords are intentionally left blank here.');
}

/**
 * Run this ONCE from the Apps Script editor to turn on the Monday-morning weekly
 * digest email (see sendWeeklyDigest_() in 04_FollowupEmails.gs). Safe to re-run --
 * deletes any existing trigger for sendWeeklyDigest_ first, so you never end up with
 * duplicate triggers firing multiple emails.
 */
function setupWeeklyDigestTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendWeeklyDigest_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendWeeklyDigest_')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(6)
    .create();
  safeAlert_('Weekly digest trigger installed: sendWeeklyDigest_ will run every Monday around 6am. Optionally set a WEEKLY_DIGEST_EMAIL script property to send it somewhere other than the office notify address.');
}

// SpreadsheetApp.getUi() only works when called from an actual open Sheets UI session (e.g.
// a custom menu item) -- it throws when a function is run directly from the Apps Script
// editor's own Run button instead, which is how every setup/migrate*_ function here is
// normally run. That's a cosmetic failure only: it happens on the confirmation popup at the
// END of each of those functions, after the real work (writing columns/headers) already
// finished. This swallows that specific failure and logs the same message to the execution
// log instead (View -> Logs, or Ctrl+Enter in the editor), so running from the editor never
// LOOKS like it failed when it actually succeeded.
function safeAlert_(msg) {
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    Logger.log(msg);
  }
}

function addMissingColumns_(tabName, cols) {
  const sheet = getSheet_(tabName);
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  const headers = headerRange.getValues()[0];
  const toAdd = cols.filter(function (col) { return headers.indexOf(col) === -1; });
  if (toAdd.length === 0) return tabName + ': nothing to do, columns already exist.';
  const startCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, startCol, 1, toAdd.length).setValues([toAdd])
    .setFontWeight('bold').setBackground('#2c4a35').setFontColor('#ffffff');
  return tabName + ': added ' + toAdd.join(', ') + '.';
}

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SCHEMA).forEach(function (tabName) {
    let sheet = ss.getSheetByName(tabName);
    if (!sheet) sheet = ss.insertSheet(tabName);
    const headers = SCHEMA[tabName];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#2c4a35').setFontColor('#ffffff');
  });
  // Remove the default empty "Sheet1" if it's still around and unused
  const def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  safeAlert_('Done -- all 6 tabs created with headers. incident_reports contains real student PII: set sharing/protection on that tab before going further.');
}

function getSheet_(tabName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) throw new Error('Sheet tab not found: ' + tabName + ' -- run setupSpreadsheet() first.');
  return sheet;
}

function getProp_(key) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  if (!v) throw new Error('Missing Script Property: ' + key + ' -- set it in Project Settings -> Script Properties.');
  return v;
}

function newUuid_() {
  return Utilities.getUuid();
}

function nowIso_() {
  return new Date().toISOString();
}

// --- site_config key/value helpers (upsert/delete -- genericUpdate_ in 03_Router.gs throws
// if no row matches, so config values that may or may not exist yet need their own helper) ---
function getSiteConfigValue_(key, fallback) {
  const sheet = getSheet_(TABS.SITE_CONFIG);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(key)) {
      const v = data[i][1];
      return (v === '' || v === null || v === undefined) ? fallback : v;
    }
  }
  return fallback;
}

function setSiteConfigValue_(key, value) {
  const sheet = getSheet_(TABS.SITE_CONFIG);
  const data = sheet.getDataRange().getValues();
  const now = nowIso_();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(key)) {
      sheet.getRange(i + 1, 1, 1, 3).setValues([[key, value, now]]);
      return;
    }
  }
  sheet.appendRow([key, value, now]);
}

function deleteSiteConfigValue_(key) {
  const sheet = getSheet_(TABS.SITE_CONFIG);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(key)) sheet.deleteRow(i + 1);
  }
}
