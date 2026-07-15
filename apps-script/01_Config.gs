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
  ],
  [TABS.INCIDENT_REPORTS]: [
    'id', 'created_at', 'student_name', 'grade', 'homeroom',
    'what_happened', 'who_was_involved', 'student_perspective',
    'status', 'assigned_to', 'staff_notes', 'follow_up_action', 'closed_at', 'closed_by',
  ],
  [TABS.SITE_CONFIG]: ['key', 'value', 'updated_at'],
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
  SpreadsheetApp.getUi().alert(results.join('\n'));
}

/**
 * Run this ONCE, only if you already ran setupSpreadsheet() before location/peak were
 * added to the reflections schema above. Same safe append-only behavior as
 * migrateAddNameGradeColumns_().
 */
function migrateAddLocationPeakColumns_() {
  const result = addMissingColumns_(TABS.REFLECTIONS, ['location', 'peak']);
  SpreadsheetApp.getUi().alert(result);
}

/**
 * Run this ONCE, only if you already ran setupSpreadsheet() before the Camp Mountaineer
 * pre-page columns were added to the reflections schema above. Same safe append-only
 * behavior as the other migrate*_ helpers.
 */
function migrateAddCampMountaineerColumns_() {
  const result = addMissingColumns_(TABS.REFLECTIONS, ['cm_reason', 'cm_initiated_by', 'cm_why', 'cm_effort_agreed', 'cm_checklist_confirmed']);
  SpreadsheetApp.getUi().alert(result);
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

  SpreadsheetApp.getUi().alert('Done -- all 6 tabs created with headers. incident_reports contains real student PII: set sharing/protection on that tab before going further.');
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
