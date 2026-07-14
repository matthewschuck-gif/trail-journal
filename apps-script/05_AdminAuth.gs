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

// Table names that require a valid staff password on every request.
const ADMIN_ONLY_TABLES_ = ['panel_sessions', 'followup_records', 'incident_reports'];

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
 */
function enforceAdminGate_(body) {
  const needsAdmin =
    (ADMIN_ONLY_TABLES_.indexOf(body.table) !== -1) ||
    (body.action === 'ai' && ADMIN_ONLY_AI_TYPES_.indexOf(body.type) !== -1) ||
    (body.action === 'sendFollowupEmails') ||
    // reflections/responder_reflections are open for insert (students submitting), but
    // reading them back in bulk (the admin dashboard listing/export) is staff-only.
    ((body.table === 'reflections' || body.table === 'responder_reflections') && body.action === 'query' && !body.filters);

  if (!needsAdmin) return null;
  return requireStaffPassword_(body.adminPassword);
}
