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

  const studentLabel = row.student_name || row.initial || 'Student';
  const today = Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd');

  // Column order matches that sheet's own header row exactly:
  // id, name, type, date, note, multiplier, BLACK, GOLD, GREY, PURPLE
  // id is left blank -- that sheet's own automation fills it in, same as manual entry.
  const rowValues = [
    '', 'Trail Journal Growth — ' + studentLabel, 'trail_journal', today,
    'High insight reflection, staff-approved — Trail Journal', '',
    squad === 'BLACK' ? 1 : '', squad === 'GOLD' ? 1 : '', squad === 'GREY' ? 1 : '', squad === 'PURPLE' ? 1 : '',
  ];
  sheet.appendRow(rowValues);

  genericUpdate_(TABS.REFLECTIONS, { session_id: sessionId }, { clash_flagged: true, clash_points_awarded: true });

  return { ok: true, squad: squad, studentLabel: studentLabel };
}
