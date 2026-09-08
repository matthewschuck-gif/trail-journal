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
const ADMIN_ONLY_TABLES_ = ['panel_sessions', 'incident_reports'];

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
  const sheet = getSheet_(TABS.STAFF_USERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0];
    const pw = data[i][1];
    if (pw !== '' && pw !== null && String(pw) === String(password)) {
      return { ok: true, name: name || 'Staff' };
    }
  }
  return { ok: false };
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
function enforceAdminGate_(body) {
  const needsAdmin =
    (ADMIN_ONLY_TABLES_.indexOf(body.table) !== -1) ||
    (body.action === 'ai' && ADMIN_ONLY_AI_TYPES_.indexOf(body.type) !== -1) ||
    (body.action === 'sendFollowupEmails' && !(body.payload && body.payload.singleEvent)) ||
    (body.action === 'sendSummaryEmail') ||
    // reflections/responder_reflections are open for insert (students submitting), but
    // reading them back in bulk (the admin dashboard listing/export) is staff-only.
    ((body.table === 'reflections' || body.table === 'responder_reflections') && body.action === 'query' && !body.filters) ||
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
    (body.table === 'site_config' && (body.action === 'insert' || body.action === 'update'));

  if (!needsAdmin) return null;
  return requireStaffPassword_(body.adminPassword);
}
