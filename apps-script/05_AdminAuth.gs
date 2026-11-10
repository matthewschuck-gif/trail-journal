/**
 * Trail Journal -- Apps Script backend
 * File 5 of N (revised again): staff password gate, checked server-side.
 *
 * The original two-deployment (public + domain-restricted) design does NOT work in practice:
 * Apps Script's "Anyone within domain" Web App deployments route requests through an extra
 * Google auth-check hop that doesn't return CORS headers, so a cross-origin fetch() from
 * trailjournal.org can never read the response, even when correctly signed in. Confirmed by
 * testing. So there is now only ONE deployment (the public one), used by every page.
 *
 * Staff-only actions are instead gated by a password checked HERE, server-side, in Apps
 * Script -- never exposed in frontend source. This is still a real improvement over the
 * original app, which compared a hardcoded password directly in client-side JS (visible to
 * anyone who viewed page source). It's a shared-secret model, similar in spirit to the
 * APP_TOKEN check in 03_Router.gs, just with its own separate secret and its own error message.
 *
 * Add this Script Property (Project Settings > Script Properties):
 *   ADMIN_PASSWORD = <a password you choose for staff -- can be the old one or a new one>
 */

// Table names that require a valid staff password on every request, for every action
// (insert/update/query alike). followup_records is deliberately NOT in this list -- see the
// carve-out below enforceAdminGate_ for why.
const ADMIN_ONLY_TABLES_ = ['panel_sessions', 'incident_reports', 'clash_requests'];

// Staff-only actions added with the Trailback pipeline (11_Trailback.gs, 07_ClashOfClasses.gs).
// approvePanelPlan also needs a Google sign-in, not just the password (checked in its function).
const STAFF_ACTIONS_ = ['followupDates', 'approvePanelPlan', 'submitPanelSession', 'clashRequests', 'decideClashRequest'];

// AI proxy "type" values that are staff-only tools, not part of the student-facing flow.
const ADMIN_ONLY_AI_TYPES_ = ['pattern_analysis', 'parent_letter', 'panel_inquiry_analysis', 'panel_recommendation_draft'];

function requireStaffPassword_(providedPassword) {
  const expected = getProp_('ADMIN_PASSWORD');
  if (!providedPassword || providedPassword !== expected) {
    throw new Error('Incorrect staff password.');
  }
  return { ok: true };
}

/**
 * Looks up a password against the staff_users sheet (name + password per row, see
 * migrateAddStaffUsersTab_() in 01_Config.gs) and returns which staff member it belongs to.
 * Powers the two mid-journal check-in gates and the final screen unlock -- each of those
 * three uses its OWN password lookup here instead of the shared ADMIN_PASSWORD, so a gate
 * unlock can be attributed to a specific person. Deliberately separate from
 * requireStaffPassword_()/ADMIN_PASSWORD above, which still gates the admin dashboard --
 * two different secrets on purpose, since this list is meant to be handed out more widely
 * (any adult who might supervise a check-in) than dashboard access should be. Not routed
 * through enforceAdminGate_ below -- this function IS the gate for its own action.
 */
function checkStaffPassword_(password) {
  if (!password) return { ok: false };
  // Slow down guessing: after 20 wrong tries in 10 minutes, every gate waits.
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('gate-fails')) || 0;
  if (fails >= 20) return { ok: false, error: 'Too many wrong passwords. Wait a few minutes.' };
  const sheet = getSheet_(TABS.STAFF_USERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0];
    const pw = data[i][1];
    if (pw !== '' && pw !== null && gatePasswordMatches_(String(pw), String(password))) {
      return { ok: true, name: name || 'Staff' };
    }
  }
  cache.put('gate-fails', String(fails + 1), 600);
  return { ok: false };
}

// ── Scrambled gate passwords ──
// A cell holds either a plain password (just typed) or "sha256$salt$hash" (scrambled by
// scrambleGatePasswords). Both work, so typing a new plain password in a cell works at once;
// the next scramble run hides it.
function gatePasswordMatches_(stored, typed) {
  const m = stored.match(/^sha256\$([^$]+)\$([0-9a-f]{64})$/);
  return m ? gateHash_(typed, m[1]) === m[2] : stored === typed;
}

function gateHash_(pw, salt) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}

/**
 * Run from the editor (or the Trail Journal menu) after adding or changing any password in
 * the staff_users tab: replaces each plain password with its scrambled form. The same
 * passwords keep working at the gates; nobody can read them back from the Sheet.
 */
function scrambleGatePasswords() {
  const sheet = getSheet_(TABS.STAFF_USERS);
  const data = sheet.getDataRange().getValues();
  let n = 0;
  for (let i = 1; i < data.length; i++) {
    const pw = data[i][1];
    if (pw === '' || pw === null || /^sha256\$/.test(String(pw))) continue;
    const salt = Utilities.getUuid().slice(0, 8);
    sheet.getRange(i + 1, 2).setValue('sha256$' + salt + '$' + gateHash_(String(pw), salt));
    n++;
  }
  safeAlert_(n ? 'Scrambled ' + n + ' gate password(s). They still work the same at the gates.' : 'No plain passwords to scramble.');
  return n;
}

/**
 * Called from the router for every request. Throws if this action needs the staff password
 * and the request didn't include a correct one. Silently passes through for anonymous
 * student-facing actions (journal + responder form inserts, and student-facing AI types).
 *
 * IMPORTANT: the Trail Journal's "Schedule Follow-Up Check-In" step runs on the STUDENT'S
 * OWN DEVICE with nobody signed in as staff -- there is no password prompt anywhere in that
 * flow, by design (Matt: "i don't want the teacher to have to log in ... has to come
 * automatically on the final page"). Two actions it calls must therefore stay open:
 *   - insert/update on followup_records (saving the record, marking the email sent/failed)
 *   - sendFollowupEmails when payload.singleEvent is true (the one custom-dated invite the
 *     student's own form builds -- as opposed to the staff Trailback Panel's batch 3-interval
 *     reminder run, which never sets singleEvent and stays fully gated below)
 * Bulk read-back of followup_records (a 'query') stays staff-only, same as reflections.
 *
 * The admin dashboard's "mark this follow-up complete" action (page-followup's Trail
 * Journal Follow-Ups section) ALSO calls update on followup_records, but that one is a real
 * staff-only action -- it's gated separately below by checking for the completion fields in
 * the patch, rather than widening the table-level exemption above (which must stay narrow
 * so it keeps covering only the student's own zero-login calendar_status update).
 */
// The student page's prior-reflections check: reflections, filtered by { student_id } (the 800#,
// once students type it) or { initial }, and nothing else.
function isStudentHistoryQuery_(body) {
  const keys = Object.keys(body.filters || {});
  return body.table === 'reflections' && body.action === 'query' && keys.length === 1 &&
    (keys[0] === 'initial' || keys[0] === 'student_id') && String(body.filters[keys[0]] || '').length > 0;
}

// True when the request comes from signed-in staff (tjCall) or carries the correct staff
// password (never throws).
function hasStaffPassword_(body) {
  if (body._staffEmail) return true;
  try { requireStaffPassword_(body.adminPassword); return true; } catch (e) { return false; }
}

/**
 * The generic insert/update/query actions a page may use WITHOUT staff access -- exactly what
 * the student journal and the responder form do, nothing more. Anything else on those actions
 * (staff_users, incident_reports, panel_sessions, unknown tabs, an update matched on anything
 * but the row's own id) needs staff. Added 2026-10-31: before this, any tab not on
 * ADMIN_ONLY_TABLES_ -- including staff_users with every adult's gate password -- could be read
 * by anyone holding the public APP_TOKEN, and an update matched on e.g. { initial } could
 * overwrite many students' rows at once.
 */
function isOpenGenericAction_(body) {
  const a = body.action, t = body.table;
  const keys = function (o) { return Object.keys(o || {}); };
  if (a === 'insert') return ['reflections', 'responder_reflections', 'followup_records'].indexOf(t) > -1;
  if (a === 'update') {
    const m = keys(body.match);
    if (t === 'reflections' || t === 'responder_reflections') return m.length === 1 && m[0] === 'session_id' && !!body.match.session_id;
    if (t === 'followup_records') return m.length === 1 && m[0] === 'id' && !!body.match.id;
    return false;
  }
  if (a === 'query') return t === 'site_config' || isStudentHistoryQuery_(body);
  return false;
}

function enforceAdminGate_(body) {
  const generic = body.action === 'insert' || body.action === 'update' || body.action === 'query';
  const needsAdmin =
    (generic && !isOpenGenericAction_(body)) ||
    (ADMIN_ONLY_TABLES_.indexOf(body.table) !== -1) ||
    (body.action === 'ai' && ADMIN_ONLY_AI_TYPES_.indexOf(body.type) !== -1) ||
    (body.action === 'sendFollowupEmails' && !(body.payload && body.payload.singleEvent)) ||
    (body.action === 'sendSummaryEmail') ||
    // reflections/responder_reflections are open for insert (students submitting), but every
    // read is staff-only. The one exception is the student page's "prior reflections" check,
    // filtered by { initial } only, which gets back just insight level, trait and date
    // (studentHistory_ in 03_Router.gs). Before 2026-10-31 any filter at all (even {} or a
    // made-up column) skipped this gate and returned every reflection in full.
    ((body.table === 'reflections' || body.table === 'responder_reflections') && body.action === 'query' && !isStudentHistoryQuery_(body)) ||
    // followup_records: bulk read-back only, not the student's own insert/update (see note above).
    (body.table === 'followup_records' && body.action === 'query') ||
    // followup_records: the staff "mark complete" action, identified by its patch touching
    // the completion columns -- gated even though table+action alone (update) is otherwise
    // exempt for the student's own calendar_status update.
    (body.table === 'followup_records' && body.action === 'update' && body.patch && Object.prototype.hasOwnProperty.call(body.patch, 'completed')) ||
    // Peaks poster uploads/resets (06_PeaksPosters.gs) are a staff-only tool. Reads of
    // site_config stay open (the student page needs to fetch overrides with no login), but
    // writes to it are gated here too, as a backstop in case anything ever calls the generic
    // insert/update actions on that table directly instead of going through those functions.
    (body.action === 'uploadPeaksPoster' || body.action === 'resetPeaksPoster') ||
    (body.action === 'setOfficeRecordRecipients') ||
    (body.action === 'awardClashPoint') ||
    (STAFF_ACTIONS_.indexOf(body.action) !== -1) ||
    (body.action === 'overrideInsightLevelHigh') ||
    (body.table === 'site_config' && (body.action === 'insert' || body.action === 'update'));

  if (!needsAdmin) return null;
  if (body._staffEmail) return { ok: true }; // signed in on the staff link (tjCall, 08_Web.gs)
  return requireStaffPassword_(body.adminPassword);
}
