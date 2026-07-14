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
 *      RESEND_API_KEY      = re_...              (existing Resend key from Supabase, or a new one)
 *      FROM_EMAIL           = noreply@trailjournal.org  (or your preferred sender)
 *      APP_TOKEN            = <any random string you make up> (shared token, see 03_Router.gs)
 *      ADMIN_DOMAIN         = yourschooldomain.org (checked against staff sign-in, see 05_AdminAuth.gs)
 * 5. Run `setupSpreadsheet` once from the editor (select it in the function dropdown, click Run).
 *    Approve the permission prompts. This creates all 6 tabs with correct header rows.
 * 6. Deploy this TWICE as separate Web App deployments (see 05_AdminAuth.gs for exactly why):
 *      Deployment 1 "public": Execute as "Me", Who has access "Anyone"
 *        -> use this URL in index.html and respond/index.html
 *      Deployment 2 "admin": Execute as "Me", Who has access "Anyone within <your domain>"
 *        -> use this URL in admin/index.html, panel/index.html, panel/followup/index.html
 *    No Google Cloud Console / OAuth client needed for either -- both are plain Apps Script
 *    deployment settings.
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
