/**
 * Trail Journal -- Apps Script backend
 * File 22: Trail Guide points. Student panelists (Trail Guides) who sit on a Summit earn their
 * own squad TRAILBACK_GUIDE_POINTS (10) each.
 *
 *   - Trail Guides are the Panel Members rows with Role "Student", Active, and an 800# and Squad
 *     (setup adds those two columns to the Panel Members tab).
 *   - The Summit form on the check-in page lists them; tick who sat on it.
 *   - Each ticked guide becomes a Clash request (Admin › Clash requests). Points reach Clash only
 *     when staff approve it there, like every Trailback point. Once per guide per case.
 *   - The student who reached the Summit gets the Redemption points (16_Summit.gs) as before.
 * TRAILBACK_GUIDE_POINTS = 0 turns this off.
 */

function guidePoints_() { const t = String(webSetting_('TRAILBACK_GUIDE_POINTS')).trim(); return t === '' ? 10 : Math.max(0, Number(t) || 0); }

/** Active Trail Guides: [{ id, name, squad }] (Panel Members, Role "Student", with 800# and Squad). */
function trailGuideList_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PANEL_MEMBERS_TAB_);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getDataRange().getValues();
  const h = rows[0].map(function (x) { return String(x).trim().toLowerCase(); });
  const col = function (name) { return h.indexOf(name); };
  const idCol = col('800#') > -1 ? col('800#') : col('student id');
  return rows.slice(1).filter(function (r) {
    const active = col('active') < 0 ? '' : String(r[col('active')]).toLowerCase();
    return String(r[col('role')] || '').toLowerCase().indexOf('student') > -1 && active !== 'false' && active !== 'no';
  }).map(function (r) {
    const email = String(r[col('email')] || '').trim();
    const id = String(idCol > -1 ? r[idCol] : '').replace(/\D/g, '') || (email.match(/^(\d{6,})@/) || [])[1] || '';
    return { id: id, name: String(r[col('name')] || '').trim(), squad: String(col('squad') > -1 ? r[col('squad')] : '').trim().toUpperCase() };
  }).filter(function (g) { return g.id && COC_SQUADS_.indexOf(g.squad) > -1; });
}

/** After a Summit is saved: one pending Clash request per ticked guide. Returns warnings. */
function requestGuidePoints_(c, ids, by) {
  const pts = guidePoints_();
  if (!pts || !Array.isArray(ids) || !ids.length) return [];
  const guides = trailGuideList_(), out = [];
  ensureSchemaTab_(TABS.CLASH_REQUESTS);
  ids.slice(0, 12).forEach(function (raw) {
    const id = String(raw || '').replace(/\D/g, '');
    const g = guides.filter(function (x) { return x.id === id; })[0];
    if (!g) { out.push('A ticked Trail Guide is not on the Panel Members tab with an 800# and Squad.'); return; }
    if (String(c.student_id || '') === id) return; // the climber gets Redemption points, not guide points
    const source = 'trail_guide:' + id;
    const dup = genericQuery_(TABS.CLASH_REQUESTS, { case_id: c.case_id, source: source }).filter(function (r) { return r.status !== 'declined'; })[0];
    if (dup) return;
    genericInsert_(TABS.CLASH_REQUESTS, { source: source, squad: g.squad, points: pts, note: TRAIL_GUIDE_NOTE_,
      case_id: c.case_id, status: 'pending', requested_by: by || '' });
  });
  return out;
}

const TRAIL_GUIDE_NOTE_ = 'A Trail Guide sat on a Summit panel.';

// Panel Members: add the 800# and Squad columns Trail Guides need (setup).
function ensureGuideColumns_() {
  const sh = ensurePanelMembersTab_();
  const width = Math.max(sh.getLastColumn(), 4);
  const h = sh.getRange(1, 1, 1, width).getValues()[0].map(function (x) { return String(x).trim().toLowerCase(); });
  // The tab's note sits in column F; new headers go in the first empty header cells from column E on.
  let next = 5;
  ['800#', 'Squad'].forEach(function (name) {
    if (h.indexOf(name.toLowerCase()) > -1) return;
    while (String(h[next - 1] || '').trim() && next <= width + 2) next++;
    sh.getRange(1, next).setValue(name).setFontWeight('bold').setBackground('#2c4a35').setFontColor('#ffffff');
    h[next - 1] = name.toLowerCase();
  });
}
