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
    (body.table === 'followup_records' && body.action === 'query');

  if (!needsAdmin) return null;
  return requireStaffPassword_(body.adminPassword);
}
