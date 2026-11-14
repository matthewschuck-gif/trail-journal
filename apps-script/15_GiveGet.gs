/**
 * Trail Journal -- Apps Script backend
 * File 15: the Action Ideas bank and the building's Give / Get board.
 *
 * ACTION IDEAS tab (staff edit any time; the bank grows): restorative actions the panel can pick.
 *   Show | Menu | Consequence Type | Idea | Done When | Need | Notes
 *   Menu: Repair, Reconnect, Reflect, Re-engage (the v2 deliberation menu)
 *   Done When: a suggested written test, so the plan's "how we know it is done" starts filled in
 *   Need: Connection, Awareness, Skills, Regulation (blank = any): ideas for the need come first
 * GIVE / GET tab (?page=givget on the staff link; each staff member keeps their own rows):
 *   GIVE = something a staff member can offer a student (a lunch group, mentoring, a quiet spot,
 *          a club seat, tutoring). GET = something a student can do for or with them (set up the
 *          lab, help in the library, run the morning cart, read to a younger class).
 *   ID | Staff Email | Staff Name | Kind | What | Menu | When | Spots | Active | Updated
 *   The whole building can see the board: the needs and the offers across the school. The
 *   panel's plan builder shows them as ready-made actions with that staff member as the
 *   check-in adult.
 */

const IDEAS_TAB_ = 'Action Ideas';
const IDEA_HEADERS_ = ['Show', 'Menu', 'Consequence Type', 'Idea', 'Done When', 'Need', 'Notes'];
const GIVEGET_TAB_ = 'Give Get';
const GIVEGET_HEADERS_ = ['ID', 'Staff Email', 'Staff Name', 'Kind', 'What', 'Menu', 'When', 'Spots', 'Active', 'Updated'];
const MENUS_ = ['Repair', 'Reconnect', 'Reflect', 'Re-engage'];

// [Menu, Consequence Type, Idea, Done When, Need]
const ACTION_IDEAS_SEED_ = [
  ['Repair', 'You Break It, You Fix It', 'Apology the person confirms (in person or written, with a trusted adult there)', 'The person harmed says it happened and felt real', ''],
  ['Repair', 'You Break It, You Fix It', 'Clean, fix or replace what was damaged', 'The staff member responsible for the space signs off', ''],
  ['Repair', 'Altruistic', 'Do something positive for the person harmed (their choice from a short list)', 'The person harmed picks it and says it happened', 'Connection'],
  ['Repair', 'Logical', 'Make up the class time or work that was lost', 'The teacher says the work is in', 'Skills'],
  ['Repair', 'Logical', 'Help the staff whose time was taken (lunch, custodial, office) for set days', 'The staff member signs off on each day', ''],
  ['Repair', 'Altruistic', 'Write a note of thanks to someone who helped clean up after it', 'The note is delivered', 'Awareness'],
  ['Repair', 'You Break It, You Fix It', 'Take down or correct the post, and tell the people who saw it', 'The post is gone and the check-in adult saw the follow-up', 'Awareness'],
  ['Repair', 'Logical', 'Restore the trust with a teacher: a re-entry talk plus one week of agreed check-ins', 'The teacher says the week went as agreed', 'Connection'],
  ['Reconnect', 'Community', 'Lunch group with positive peers', 'Attended the agreed number of days (lunch staff sign off)', 'Connection'],
  ['Reconnect', 'Inspire', 'Coach or mentor meeting', 'Met the mentor the agreed number of times', 'Connection'],
  ['Reconnect', 'Community', 'Try out for a team or join a club', 'Went to the first two meetings or the tryout', 'Connection'],
  ['Reconnect', 'Community', 'Join a squad job for Clash of Classes (Generals pick it with the student)', 'The squad General says the job is done', 'Connection'],
  ['Reconnect', 'Inspire', 'Shadow an adult in a job the student is curious about for a period', 'The adult signs off', 'Connection'],
  ['Reconnect', 'Community', 'Restorative circle with the people involved', 'The circle happened; the facilitator signs off', 'Awareness'],
  ['Reconnect', 'Community', 'Check in with a trusted adult every morning for a week', 'Five check-ins logged', 'Regulation'],
  ['Reconnect', 'Inspire', 'Lunchbassadors: help a new or lonely student at lunch', 'Two weeks of lunches, lunch staff sign off', 'Connection'],
  ['Reflect', 'Teach Skills', 'Poster, PSA or slideshow that teaches others what they learned', 'Shown in advisory or posted with the teacher\'s OK', 'Awareness'],
  ['Reflect', 'Plan', 'A Camp Mountaineer pathway', 'The Camp Mountaineer coordinator marks it complete', 'Skills'],
  ['Reflect', 'Plan', 'A written if-then plan for next time ("If ___ happens, I will ___")', 'The check-in adult has a copy and the student can say it back', 'Regulation'],
  ['Reflect', 'Teach Skills', 'Trail Journal follow-up reflection after two weeks', 'The follow-up journal is submitted', 'Awareness'],
  ['Reflect', 'Teach Skills', 'Interview an adult about how they handle the same kind of moment', 'The student shares three things they learned with the check-in adult', 'Skills'],
  ['Reflect', 'Plan', 'Map the pattern: when, where, with whom it happens, and one change for each', 'The map is done and reviewed with the check-in adult', 'Awareness'],
  ['Reflect', 'Teach Skills', 'Practice a calm-down tool three times and report which one worked', 'Three practices logged', 'Regulation'],
  ['Reflect', 'Teach Skills', 'Read or watch something on the topic and teach it back', 'The teach-back happened', 'Awareness'],
  ['Re-engage', 'Choice', 'Service project for the school', 'The project is finished and the sponsor signs off', ''],
  ['Re-engage', 'Transform', 'A job in the building (library, office, cafeteria, a classroom)', 'The staff member signs off on each shift', 'Connection'],
  ['Re-engage', 'Space and Time', 'Re-entry meeting with the teacher', 'The meeting happened; the teacher signs off', 'Connection'],
  ['Re-engage', 'Loss of Privilege, with a path back', 'Earn back the privilege (phone, hallway pass, seat choice) with a clear target', 'The target is met and the privilege is back', 'Regulation'],
  ['Re-engage', 'Space and Time', 'A planned break spot and signal with the teacher', 'Used appropriately for two weeks (teacher says so)', 'Regulation'],
  ['Re-engage', 'Choice', 'Pick between two tasks that help the area that was affected', 'The task the student chose is done', ''],
  ['Re-engage', 'Transform', 'Help run a squad or school event', 'The event lead signs off', 'Connection'],
  ['Re-engage', 'Space and Time', 'A new seat or schedule tweak the student helps choose', 'Two weeks with no repeat (teacher says so)', 'Regulation'],
];

function ensureGiveGetTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [[IDEAS_TAB_, IDEA_HEADERS_, ACTION_IDEAS_SEED_.map(function (r) { return [true, r[0], r[1], r[2], r[3], r[4], '']; })],
   [GIVEGET_TAB_, GIVEGET_HEADERS_, null]].forEach(function (t) {
    if (ss.getSheetByName(t[0])) return;
    const sh = ss.insertSheet(t[0]);
    sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#2c4a35').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    if (t[2]) sh.getRange(2, 1, t[2].length, t[1].length).setValues(t[2]);
  });
}

function tabRows_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getDataRange().getValues(), h = v[0].map(function (x) { return String(x).trim(); });
  return v.slice(1).map(function (r, i) { const o = { _row: i + 2 }; h.forEach(function (k, j) { o[k] = r[j]; }); return o; });
}
function ticked2_(v) { return v === true || /^(true|yes|y|x)$/i.test(String(v || '').trim()); }

function giveGetView_(r) {
  return { id: String(r['ID']), email: String(r['Staff Email'] || '').toLowerCase(), name: String(r['Staff Name'] || ''),
    kind: /^get/i.test(String(r['Kind'])) ? 'Get' : 'Give', what: String(r['What'] || ''), menu: MENUS_.indexOf(String(r['Menu'])) > -1 ? String(r['Menu']) : '',
    when: String(r['When'] || ''), spots: String(r['Spots'] || ''), active: ticked2_(r['Active']) };
}

/** The plan builder's bank: the ideas (Show ticked) and the active Give / Get offers. */
function actionIdeas_(body) {
  if (!body._staffEmail && !body._admin) throw new Error('Staff only.');
  ensureGiveGetTabs_();
  const ideas = tabRows_(IDEAS_TAB_).filter(function (r) { return ticked2_(r['Show']) && String(r['Idea'] || '').trim(); }).map(function (r) {
    return { menu: String(r['Menu'] || ''), type: String(r['Consequence Type'] || ''), idea: String(r['Idea']), test: String(r['Done When'] || ''), need: String(r['Need'] || '') };
  });
  const offers = tabRows_(GIVEGET_TAB_).map(giveGetView_).filter(function (g) { return g.active && g.what; });
  return { ideas: ideas, offers: offers };
}

/** The Give / Get page: everyone's active rows (the building's needs and offers) and my rows. */
function giveGetBoard_(body) {
  const me = String(body._staffEmail || '').toLowerCase();
  if (!me) throw new Error('Open the Give / Get page from the staff link and sign in with your school Google account.');
  ensureGiveGetTabs_();
  const all = tabRows_(GIVEGET_TAB_).map(giveGetView_);
  return { me: me, myName: panelNameFromEmail_(me), menus: MENUS_,
    mine: all.filter(function (g) { return g.email === me; }),
    board: all.filter(function (g) { return g.active && g.what; }) };
}

/** Add, change or remove one of my rows: { id?, kind, what, menu, when, spots, active, remove? }. */
function giveGetSave_(body) {
  const me = String(body._staffEmail || '').toLowerCase();
  if (!me) throw new Error('Sign in with your school Google account.');
  ensureGiveGetTabs_();
  const p = body.payload || {};
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GIVEGET_TAB_);
  const rows = tabRows_(GIVEGET_TAB_);
  const lead = panelIsAdmin_(me);
  const existing = p.id ? rows.filter(function (r) { return String(r['ID']) === String(p.id); })[0] : null;
  if (p.id && !existing) throw new Error('That row is gone.');
  if (existing && String(existing['Staff Email']).toLowerCase() !== me && !lead) throw new Error('You can only change your own rows.');
  if (existing && p.remove) { sh.deleteRow(existing._row); return { ok: true }; }
  const what = String(p.what || '').trim().slice(0, 200);
  if (!what) throw new Error('Say what you can give or what students could do.');
  const vals = { 'ID': existing ? existing['ID'] : newUuid_().slice(0, 8), 'Staff Email': existing ? existing['Staff Email'] : me,
    'Staff Name': String(p.name || (existing && existing['Staff Name']) || panelNameFromEmail_(me)).slice(0, 60),
    'Kind': /^get/i.test(String(p.kind)) ? 'Get' : 'Give', 'What': what, 'Menu': MENUS_.indexOf(String(p.menu)) > -1 ? String(p.menu) : '',
    'When': String(p.when || '').trim().slice(0, 100), 'Spots': String(p.spots || '').trim().slice(0, 20), 'Active': p.active !== false, 'Updated': nowIso_() };
  const row = GIVEGET_HEADERS_.map(function (h) { return vals[h]; });
  if (existing) sh.getRange(existing._row, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
  return { ok: true, id: String(vals['ID']) };
}

function showGiveGetLink() {
  let url = '';
  try { url = ScriptApp.getService().getUrl() + '?page=givget'; } catch (e) { url = ''; }
  safeAlert_('Give / Get for staff (school sign-in):\n\n' + url + '\n\nStaff list what they can offer students (Give) and what students can do for them (Get). ' +
    'The Trailback Panel sees these as ready-made actions. New action ideas go on the Action Ideas tab.');
}
