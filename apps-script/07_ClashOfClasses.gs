/**
 * Trail Journal -- Apps Script backend
 * File 7 of N: Clash of Classes point-award bridge.
 *
 * clashofclasses.org's real live backend is a separate Google Sheet ("Clash of Classes
 * Points", COC_SHEET_ID_ below) -- that site polls its "events" tab and reflects new rows
 * within about 5-10 seconds, per that sheet's own "How To Use This Sheet" tab. This file
 * does programmatically, on admin approval, exactly what that tab's instructions tell a
 * staff member to do by hand: append one row with a 1 in the awarded squad's column.
 *
 * This REPLACES the old approach (a client-side fetch() in index.html's now-removed
 * awardClashPoints(), PATCHing a Supabase project directly the instant a submission scored
 * High insight -- no staff involved at all). Two reasons that had to change: Matt wants a
 * human reviewing/approving before a point is given (see the Clash Flags admin page), and
 * that Supabase project is very likely no longer what clashofclasses.org actually reads --
 * the site's own backend has since moved to this Sheet, same migration Trail Journal itself
 * went through.
 *
 * DEPLOYMENT NOTE: opening a spreadsheet other than the one this project is bound to
 * requires the broader https://www.googleapis.com/auth/spreadsheets OAuth scope. The first
 * time this runs after being deployed, Google will likely show Matt an extra permission
 * consent screen (same one-time thing as the original "send email as you" authorization) --
 * that's expected, not an error.
 */
const COC_SHEET_ID_ = '1wF01Limr7uO36DRGC6qh3XfYA7dBdXEfGRWuasmEf0w';
const COC_EVENTS_TAB_ = 'events';
const COC_SQUADS_ = ['BLACK', 'GOLD', 'GREY', 'PURPLE'];

/**
 * Squad totals for the journal's leaderboard: the sum of each squad's column on the Clash
 * "events" tab. Totals only, no names. Open to students (the leaderboard needs no sign-in).
 */
function clashStandings_() {
  const totals = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
  const sh = SpreadsheetApp.openById(COC_SHEET_ID_).getSheetByName(COC_EVENTS_TAB_);
  if (!sh) return totals;
  const values = sh.getDataRange().getValues();
  const h = values[0].map(function (x) { return String(x).trim().toUpperCase(); });
  values.slice(1).forEach(function (r) {
    COC_SQUADS_.forEach(function (s) { totals[s] += Number(r[h.indexOf(s)]) || 0; });
  });
  return totals;
}

/**
 * payload: { sessionId }. Looks up the reflection by session_id (must already have a squad
 * on file -- see the 'squad' column added to SCHEMA.reflections in 01_Config.gs), appends
 * one row to the Clash of Classes Points sheet's "events" tab awarding that squad 1 point,
 * then marks clash_points_awarded (and clash_flagged, for consistency) true on the Trail
 * Journal row so it can't be double-awarded and the admin page reflects the new status.
 * Gated admin-only -- see enforceAdminGate_ in 05_AdminAuth.gs.
 */
function awardClashPoint_(payload) {
  const sessionId = payload && payload.sessionId;
  if (!sessionId) throw new Error('sessionId is required.');

  const rows = genericQuery_(TABS.REFLECTIONS, { session_id: sessionId }, 1);
  const row = rows && rows[0];
  if (!row) throw new Error('No reflection found for session ' + sessionId);
  if (row.clash_points_awarded) throw new Error('A point was already awarded for this submission.');

  const squad = String(row.squad || '').trim().toUpperCase();
  if (COC_SQUADS_.indexOf(squad) === -1) {
    throw new Error('This submission has no valid squad on file (got: ' + (row.squad || '—') + ') -- cannot award a point.');
  }

  const sheet = SpreadsheetApp.openById(COC_SHEET_ID_).getSheetByName(COC_EVENTS_TAB_);
  if (!sheet) throw new Error('Could not find the "' + COC_EVENTS_TAB_ + '" tab on the Clash of Classes Points sheet.');

  const studentLabel = row.student_name || row.initial || 'Student'; // for the staff reply only
  const today = Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd');

  // Column order matches that sheet's own header row exactly:
  // id, name, type, date, note, multiplier, BLACK, GOLD, GREY, PURPLE
  // id is left blank -- that sheet's own automation fills it in, same as manual entry.
  // No student name: the Clash sheet feeds a public site, and a Trail Journal follows an incident.
  const rowValues = [
    '', 'Trail Journal Growth', 'trail_journal', today,
    'High insight reflection, staff-approved — Trail Journal', '',
    squad === 'BLACK' ? 1 : '', squad === 'GOLD' ? 1 : '', squad === 'GREY' ? 1 : '', squad === 'PURPLE' ? 1 : '',
  ];
  sheet.appendRow(rowValues);

  genericUpdate_(TABS.REFLECTIONS, { session_id: sessionId }, { clash_flagged: true, clash_points_awarded: true });

  return { ok: true, squad: squad, studentLabel: studentLabel };
}

// ── TRAILBACK PANEL POINTS (approval queue) ─────────────────────────────
// A completed Trailback plan asks for a squad point; staff approve it on the admin Clash Flags
// page, and only then does one row go to the Clash events tab. Neither the request nor the
// row carries a name or initials -- just the squad and "Trailback plan completed."
// Settings tab: TRAILBACK_CLASH_POINTS (default 1).
const TRAILBACK_CLASH_NOTE_ = 'Trailback plan completed.';

function trailbackClashPoints_() {
  const n = Number(webSetting_('TRAILBACK_CLASH_POINTS'));
  return n > 0 ? n : 1;
}

/**
 * Internal (the case closing as Completed calls it -- 11_Trailback.gs / Part B).
 * opts: { squad, caseId, requestedBy }. Returns the request row.
 */
function requestClashPoint_(opts) {
  const squad = String((opts && opts.squad) || '').trim().toUpperCase();
  if (COC_SQUADS_.indexOf(squad) === -1) throw new Error('A Clash point needs the student\'s squad (got: ' + ((opts && opts.squad) || 'none') + ').');
  ensureSchemaTab_(TABS.CLASH_REQUESTS);
  const caseId = String((opts && opts.caseId) || '');
  if (caseId) {
    const dup = genericQuery_(TABS.CLASH_REQUESTS, { case_id: caseId }).filter(function (r) { return r.status !== 'declined'; })[0];
    if (dup) return dup;
  }
  return genericInsert_(TABS.CLASH_REQUESTS, {
    source: 'trailback_panel', squad: squad, points: trailbackClashPoints_(), note: TRAILBACK_CLASH_NOTE_,
    case_id: caseId, status: 'pending', requested_by: (opts && opts.requestedBy) || '',
  });
}

/** Staff: the queue for the admin page. Squad, points, note, date and status only. */
function clashRequests_() {
  ensureSchemaTab_(TABS.CLASH_REQUESTS);
  return genericQuery_(TABS.CLASH_REQUESTS, {}).map(function (r) {
    return { id: r.id, created_at: r.created_at, source: r.source, squad: r.squad, points: r.points, note: r.note,
      status: r.status, decided_by: r.decided_by, decided_at: r.decided_at };
  }).reverse();
}

/**
 * Staff: body.payload { id, decision: 'approve' | 'decline' }. Approving appends one row to the
 * Clash events tab: '' | Trailback Panel | trailback_panel | date | Trailback plan completed. | '' | squad points.
 */
function decideClashRequest_(body) {
  const p = body.payload || {};
  const decision = p.decision === 'decline' ? 'decline' : p.decision === 'approve' ? 'approve' : '';
  if (!p.id || !decision) throw new Error('id and decision (approve or decline) are required.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = genericQuery_(TABS.CLASH_REQUESTS, { id: p.id }, 1)[0];
    if (!req) throw new Error('No Clash request ' + p.id);
    if (req.status !== 'pending') throw new Error('This request was already ' + req.status + '.');
    const squad = String(req.squad || '').toUpperCase();
    if (decision === 'approve') {
      if (COC_SQUADS_.indexOf(squad) === -1) throw new Error('This request has no valid squad.');
      const sheet = SpreadsheetApp.openById(COC_SHEET_ID_).getSheetByName(COC_EVENTS_TAB_);
      if (!sheet) throw new Error('Could not find the "' + COC_EVENTS_TAB_ + '" tab on the Clash of Classes Points sheet.');
      const pts = Number(req.points) > 0 ? Number(req.points) : 1;
      const today = Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd');
      sheet.appendRow(['', 'Trailback Panel', 'trailback_panel', today, TRAILBACK_CLASH_NOTE_, '',
        squad === 'BLACK' ? pts : '', squad === 'GOLD' ? pts : '', squad === 'GREY' ? pts : '', squad === 'PURPLE' ? pts : '']);
    }
    const patch = { status: decision === 'approve' ? 'approved' : 'declined', decided_by: body._staffEmail || 'staff password', decided_at: nowIso_() };
    genericUpdate_(TABS.CLASH_REQUESTS, { id: p.id }, patch);
    return Object.assign({ id: p.id, squad: squad }, patch);
  } finally {
    lock.releaseLock();
  }
}
