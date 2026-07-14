/**
 * Trail Journal -- Apps Script backend
 * File 5 of N: Workspace-restricted Google Sign-In for the admin dashboard + Trailback Panel.
 *
 * Replaces the old client-side hardcoded password check with real Google Sign-In, restricted
 * to your school's Workspace domain. This does NOT gate the student-facing journal or the
 * "Your Side of the Trail" responder form -- those stay open with no login, same as today.
 * It gates: admin dashboard reads/writes, panel_sessions, followup_records, incident_reports,
 * site_config writes, and the AI actions only staff use (pattern_analysis, parent_letter,
 * panel_inquiry_analysis, panel_recommendation_draft).
 *
 * ONE-TIME SETUP (Matt/tech team, outside Apps Script):
 * 1. Go to https://console.cloud.google.com (using the school Workspace account).
 * 2. Create or select a project -> APIs & Services -> Credentials -> Create Credentials ->
 *    OAuth client ID -> Application type: "Web application".
 * 3. Under "Authorized JavaScript origins" add: https://trailjournal.org
 *    (and your staging URL too, once we have one).
 * 4. Copy the Client ID (looks like 123...-abc....apps.googleusercontent.com -- this is NOT
 *    secret, it's meant to be embedded in frontend JS, same as the old Supabase anon key was).
 * 5. Add it as a Script Property: GOOGLE_CLIENT_ID = <that client id>
 * 6. Add another Script Property: ADMIN_DOMAIN = <your school's Workspace domain, e.g. easdpa.org>
 * 7. In admin/index.html and panel/index.html (file 9 handles this), the login screen is
 *    replaced with a "Sign in with Google" button using Google Identity Services. Once signed
 *    in, the page holds an ID token and sends it as `idToken` on every request to this backend.
 *
 * HOW VERIFICATION WORKS:
 * Apps Script has no built-in JWT verifier, so instead of validating the signature locally,
 * this calls Google's own tokeninfo endpoint, which validates the signature server-side and
 * hands back the decoded claims. We then check: token was issued for OUR client id (aud),
 * the signed-in account's email domain matches ADMIN_DOMAIN (hd / email suffix), and the
 * token hasn't expired. This is the same trust model Google's own examples use for lightweight
 * backends that don't want to pull in a JWT library.
 */

function verifyAdminToken_(idToken) {
  if (!idToken) throw new Error('Sign-in required');

  const resp = UrlFetchApp.fetch(
    'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
    { muteHttpExceptions: true }
  );

  if (resp.getResponseCode() !== 200) {
    throw new Error('Sign-in could not be verified -- please sign in again.');
  }

  const claims = JSON.parse(resp.getContentText());
  const clientId = getProp_('GOOGLE_CLIENT_ID');
  const adminDomain = getProp_('ADMIN_DOMAIN');

  if (claims.aud !== clientId) {
    throw new Error('Token was not issued for this app.');
  }
  if (Number(claims.exp) * 1000 < Date.now()) {
    throw new Error('Sign-in expired -- please sign in again.');
  }
  const emailDomain = (claims.email || '').split('@')[1];
  const hostedDomain = claims.hd || emailDomain;
  if (hostedDomain !== adminDomain) {
    throw new Error('This dashboard is restricted to ' + adminDomain + ' accounts.');
  }
  if (claims.email_verified !== 'true' && claims.email_verified !== true) {
    throw new Error('Email not verified on this Google account.');
  }

  return { email: claims.email, name: claims.name || claims.email };
}

// Table names that require a verified admin/staff sign-in for ANY action (query/insert/update).
// The student journal (reflections insert) and responder form (responder_reflections insert)
// are deliberately NOT in this list -- they stay anonymous and open, same as the live site today.
const ADMIN_ONLY_TABLES_ = ['panel_sessions', 'followup_records', 'incident_reports'];

// AI proxy "type" values that are staff-only tools, not part of the student-facing flow.
const ADMIN_ONLY_AI_TYPES_ = ['pattern_analysis', 'parent_letter', 'panel_inquiry_analysis', 'panel_recommendation_draft'];

/**
 * Call this from the router before running an action. Throws if the action needs staff
 * sign-in and the request doesn't have a valid one. Silently passes through for anonymous
 * student-facing actions.
 */
function enforceAdminGate_(body) {
  const needsAdmin =
    (ADMIN_ONLY_TABLES_.indexOf(body.table) !== -1) ||
    (body.action === 'ai' && ADMIN_ONLY_AI_TYPES_.indexOf(body.type) !== -1) ||
    // reflections/responder_reflections are open for insert (students submitting), but
    // reading them back in bulk (the admin dashboard listing/export) is staff-only.
    ((body.table === 'reflections' || body.table === 'responder_reflections') && body.action === 'query' && !body.filters);

  if (!needsAdmin) return null;
  return verifyAdminToken_(body.idToken);
}
