// Offline checks for the Trail Journal backend's access rules (fake Google services, real .gs).
//   node tests/backend.test.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const dir = path.join(__dirname, '..', 'apps-script');
let passed = 0;
function test(name, fn) { fn(); passed++; console.log('ok - ' + name); }
const J = x => JSON.parse(JSON.stringify(x));

function world() {
  const user = { email: '' };
  const ctx = { console, JSON, Math, Date, Object, Array, String, Number,
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: () => null }) }, // no Settings tab: defaults
    CacheService: { getScriptCache: function () { const c = this.__c || (this.__c = {}); return { get: k => c[k] || null, put: (k, v) => { c[k] = v; }, remove: k => { delete c[k]; } }; } },
    Utilities: { getUuid: () => require('crypto').randomUUID() },
    Session: { getActiveUser: () => ({ getEmail: () => user.email }) } };
  ctx.user = user;
  vm.createContext(ctx);
  vm.runInContext(fs.readdirSync(dir).filter(f => f.endsWith('.gs')).sort().map(f => fs.readFileSync(path.join(dir, f), 'utf8')).join('\n') + ';this.SCHEMA = SCHEMA;', ctx);
  const headers = ctx.SCHEMA.reflections;
  const row = o => headers.map(h => o[h] === undefined ? '' : o[h]);
  const reflections = [headers,
    row({ session_id: 's1', initial: 'A', student_name: 'Ann Lee', insight_level: 'High', created_at: '2026-10-01', part3_json: JSON.stringify({ trait: 'Hope' }), ai_staff_summary: 'private note' }),
    row({ session_id: 's2', initial: 'B', student_name: 'Bo Kim', insight_level: 'Low', created_at: '2026-10-02' })];
  const sheets = { reflections };
  ctx.getSheet_ = t => ({ getDataRange: () => ({ getValues: () => (sheets[t] || [ctx.SCHEMA[t] || []]).map(r => r.slice()) }) });
  ctx.getProp_ = k => ({ APP_TOKEN: 'tok', ADMIN_PASSWORD: 'staffpw' })[k];
  let out;
  ctx.jsonOut_ = o => { out = o; return o; };
  ctx.call = body => { ctx.handleRequest_({ postData: { contents: JSON.stringify(Object.assign({ token: 'tok' }, body)) } }, 'POST'); return J(out); };
  return ctx;
}

test('reading reflections with an empty or made-up filter needs the staff password (was wide open)', () => {
  const w = world();
  for (const filters of [{}, { nothing: 1 }, { student_name: 'Ann Lee' }]) {
    const r = w.call({ action: 'query', table: 'reflections', filters });
    assert.ok(r.error && /staff password/i.test(r.error), JSON.stringify(filters) + ' → ' + JSON.stringify(r));
  }
  const ok = w.call({ action: 'query', table: 'reflections', filters: {}, adminPassword: 'staffpw' });
  assert.strictEqual((ok.data || ok.result || ok).length || ok.data.length, 2);
});

test('the student prior-reflections check gets only insight level, trait and date', () => {
  const w = world();
  const r = w.call({ action: 'query', table: 'reflections', filters: { initial: 'A' } });
  const rows = r.data || r.result;
  assert.deepStrictEqual(rows, [{ insight_level: 'High', created_at: '2026-10-01', part3_json: { trait: 'Hope' } }]);
  assert.ok(!JSON.stringify(r).includes('Ann Lee') && !JSON.stringify(r).includes('private note'));
});

test('an unknown filter column matches nothing for staff too', () => {
  const w = world();
  assert.deepStrictEqual(J(w.genericQuery_('reflections', { not_a_column: 'x' })), []);
  assert.strictEqual(w.genericQuery_('reflections', { session_id: 's2' })[0].student_name, 'Bo Kim');
});

test('responder reflections: every read needs the staff password', () => {
  const w = world();
  const r = w.call({ action: 'query', table: 'responder_reflections', filters: { initial: 'A' } });
  assert.ok(r.error && /staff password/i.test(r.error));
});

test('the gate-password list and other staff tabs cannot be read or changed without staff access', () => {
  const w = world();
  for (const body of [
    { action: 'query', table: 'staff_users', filters: {} },
    { action: 'insert', table: 'staff_users', row: { name: 'x', password: 'y' } },
    { action: 'query', table: 'incident_reports', filters: {} },
    { action: 'update', table: 'reflections', match: { initial: 'A' }, patch: { status: 'x' } },
    { action: 'update', table: 'followup_records', match: {}, patch: {} },
    { action: 'insert', table: 'site_config', row: {} },
  ]) {
    const r = w.call(body);
    assert.ok(r.error && /staff password/i.test(r.error), JSON.stringify(body) + ' → ' + JSON.stringify(r));
  }
});

test('the student journal and responder form still work without staff access', () => {
  const w = world();
  w.genericInsert_ = () => ({ inserted: true });
  w.genericUpdate_ = () => ({ updated: 1 });
  for (const body of [
    { action: 'query', table: 'site_config', filters: {} },
    { action: 'insert', table: 'reflections', row: { session_id: 's9' } },
    { action: 'update', table: 'reflections', match: { session_id: 's9' }, patch: { status: 'complete' } },
    { action: 'insert', table: 'followup_records', row: {} },
    { action: 'update', table: 'followup_records', match: { id: 'f1' }, patch: { calendar_status: {} } },
    { action: 'insert', table: 'responder_reflections', row: {} },
  ]) {
    w.getSheet_ = () => ({ getDataRange: () => ({ getValues: () => [w.SCHEMA[body.table] || []] }) });
    const r = w.call(body);
    assert.ok(r.ok, JSON.stringify(body) + ' → ' + JSON.stringify(r));
  }
});

test('gate passwords: scrambled in the Sheet, the same passwords still unlock, per adult', () => {
  const w = world();
  const cache = {};
  w.CacheService = { getScriptCache: () => ({ get: k => cache[k] || null, put: (k, v) => { cache[k] = v; } }) };
  w.Utilities = { getUuid: () => 'abcdef1234', DigestAlgorithm: {}, Charset: {},
    computeDigest: (a, s) => Array.from(require('crypto').createHash('sha256').update(s).digest()).map(b => b > 127 ? b - 256 : b) };
  w.safeAlert_ = () => {};
  const staff = [['name', 'password'], ['Mr. Schuck', 'trail1'], ['Mrs. Rigg', 'peak2'], ['Guest', '']];
  w.getSheet_ = () => ({ getDataRange: () => ({ getValues: () => staff.map(r => r.slice()) }), getRange: (r, c) => ({ setValue: v => { staff[r - 1][c - 1] = v; } }) });
  assert.strictEqual(w.scrambleGatePasswords(), 2);
  assert.ok(staff.slice(1, 3).every(r => /^sha256\$/.test(r[1]) && !/trail1|peak2/.test(r[1])));
  assert.deepStrictEqual(J(w.checkStaffPassword_('peak2')), { ok: true, name: 'Mrs. Rigg' }, 'still says who unlocked it');
  assert.strictEqual(w.checkStaffPassword_('nope').ok, false);
  staff[3][1] = 'newpw'; // a new plain password works right away
  assert.strictEqual(w.checkStaffPassword_('newpw').name, 'Guest');
  assert.strictEqual(w.scrambleGatePasswords(), 1, 'only the new one is scrambled');
});

test('pages route (tjCall): school staff sign-in counts as staff; students and fake flags do not', () => {
  const w = world();
  const call = b => JSON.parse(w.tjCall(JSON.stringify(b)));
  // A student (or anyone not at the staff domain) can't read reflections, even claiming to be staff.
  w.user.email = '8001111@gomounts.org';
  let r = call({ action: 'query', table: 'reflections', filters: {}, _staffEmail: 'boss@easdpa.org' });
  assert.ok(!r.ok && /staff password/i.test(r.error), JSON.stringify(r));
  // Signed-in staff on the staff link: no password needed, full rows.
  w.user.email = 'teacher@easdpa.org';
  r = call({ action: 'query', table: 'reflections', filters: {} });
  assert.ok(r.ok && r.data.length === 2);
  // The old token route never trusts a _staffEmail in the request.
  const old = w.call({ action: 'query', table: 'reflections', filters: {}, _staffEmail: 'teacher@easdpa.org' });
  assert.ok(!old.ok && /staff password/i.test(old.error));
});

test('prior-reflections check by 800# when the student typed it', () => {
  const w = world();
  w.SCHEMA.reflections.indexOf('student_id') > -1 || assert.fail('student_id column');
  const r = w.call({ action: 'query', table: 'reflections', filters: { student_id: '800-1111' } });
  assert.ok(r.ok, JSON.stringify(r));
  assert.deepStrictEqual(r.data, [], 'no reflection has that 800# yet');
  assert.ok(!w.call({ action: 'query', table: 'reflections', filters: { student_id: '8001111', initial: 'A' } }).ok, 'only one filter allowed');
});

test('journal leaderboard: squad totals from the Clash events tab, open to students', () => {
  const w = world();
  w.SpreadsheetApp = Object.assign({}, w.SpreadsheetApp, { openById: () => ({ getSheetByName: () => ({ getDataRange: () => ({ getValues: () => [
    ['id', 'name', 'type', 'date', 'note', 'multiplier', 'BLACK', 'GOLD', 'GREY', 'PURPLE'], [1, 'a', 'x', '', '', '', 5, '', 2, 0], [2, 'b', 'x', '', '', '', 1, 4, '', ''],
  ] }) }) }) });
  const r = w.call({ action: 'clashStandings' });
  assert.ok(r.ok, JSON.stringify(r));
  assert.deepStrictEqual(r.data, { BLACK: 6, GOLD: 4, GREY: 2, PURPLE: 0 });
});

test('Clash point for a high-insight journal: the row sent to the public Clash sheet has no student name', () => {
  const w = world();
  const sent = [];
  w.Utilities = { formatDate: () => '2026-11-01' };
  w.SpreadsheetApp = Object.assign({}, w.SpreadsheetApp, { openById: () => ({ getSheetByName: () => ({ appendRow: row => sent.push(row) }) }) });
  w.genericQuery_ = () => [{ squad: 'GOLD', student_name: 'Jamie Doe', clash_points_awarded: false }];
  w.genericUpdate_ = () => {};
  const r = J(w.awardClashPoint_({ sessionId: 's1' }));
  assert.deepStrictEqual([r.ok, r.squad, r.studentLabel], [true, 'GOLD', 'Jamie Doe'], 'staff still see who it was');
  assert.strictEqual(sent.length, 1);
  assert.deepStrictEqual([sent[0][1], sent[0][2]], ['Trail Journal Growth', 'trail_journal']);
  assert.ok(!JSON.stringify(sent[0]).includes('Jamie'), 'no name in the Clash row');
  assert.strictEqual(sent[0][7], 1, 'GOLD column');
});


test('staff sessions: the password is traded for a token once; tokens work, end on sign-out, and wrong tries lock out', () => {
  const w = world();
  assert.throws(() => w.tjStaffSignIn('nope'), /Incorrect/);
  const tok = w.tjStaffSignIn('staffpw');
  assert.ok(/^tjs_[A-Za-z0-9]{40,}$/.test(tok), 'a long random token, not the password');
  const ok = w.call({ action: 'query', table: 'reflections', filters: {}, adminPassword: tok });
  assert.ok(!ok.error, JSON.stringify(ok));
  assert.ok(w.call({ action: 'query', table: 'reflections', filters: {}, adminPassword: 'tjs_' + 'a'.repeat(64) }).error, 'a made-up token fails');
  w.tjStaffSignOut(tok);
  assert.ok(/staff/i.test(w.call({ action: 'query', table: 'reflections', filters: {}, adminPassword: tok }).error), 'signed out on the server');
  for (let i = 0; i < 10; i++) { try { w.tjStaffSignIn('wrong' + i); } catch (e) { /* counted */ } }
  assert.throws(() => w.tjStaffSignIn('staffpw'), /Too many/, 'locked for 10 minutes after 10 wrong tries');
  const w2 = world();
  for (let i = 0; i < 30; i++) w2.call({ action: 'query', table: 'reflections', filters: { initial: 'A' } });
  assert.ok(w2.tjStaffSignIn('staffpw'), 'student calls with no password never lock staff out');
});

test('Staff Access (21_StaffAccess.gs): after setup, the admin tools need the list; any school account still counts for sign-up; password fallback, OFF switch, Staff Actions, alerts', () => {
  const w = world();
  const call = b => JSON.parse(w.tjCall(JSON.stringify(b)));
  const tabs = {};
  const mkSheet = (name, rows) => {
    const s = { rows, getName: () => name, getLastRow: () => s.rows.length, getMaxRows: () => 50,
      getDataRange: () => ({ getValues: () => s.rows.map(r => r.slice()) }),
      getRange: (r, c, nr, nc) => {
        const api = { getValues: () => s.rows.slice(r - 1, r - 1 + (nr || 1)).map(x => { const o = x.slice(c - 1, c - 1 + (nc || 1)); while (o.length < (nc || 1)) o.push(''); return o; }),
          setValues: v => { v.forEach((row, i) => { s.rows[r - 1 + i] = s.rows[r - 1 + i] || []; row.forEach((x, j) => { s.rows[r - 1 + i][c - 1 + j] = x; }); }); return px; },
          setValue: v => api.setValues([[v]]) };
        const px = new Proxy(api, { get: (t, p) => p in t ? t[p] : () => px });
        return px;
      },
      appendRow: r => s.rows.push(r), setFrozenRows() {} };
    tabs[name] = s; return s;
  };
  mkSheet('Staff Access', [['Email', 'Name', 'Role', 'Active', 'Last Sign-in'], ['lead@easdpa.org', 'Ms. Lead', 'Counselor', true, '']]);
  const mails = [];
  w.SpreadsheetApp.getActiveSpreadsheet = () => ({ getSheetByName: n => tabs[n] || null, insertSheet: n => mkSheet(n, []), getId: () => 'TJ' });
  w.Session.getEffectiveUser = () => ({ getEmail: () => 'owner@easdpa.org' });
  w.Session.getScriptTimeZone = () => 'America/New_York';
  w.Utilities.formatDate = d => d.toISOString().slice(0, 16).replace('T', ' ');
  w.MailApp = { sendEmail: o => mails.push(o) };
  w.PropertiesService = { getScriptProperties: () => { const p = w.__p || (w.__p = {}); return { getProperty: k => p[k] || null, setProperty: (k, v) => { p[k] = v; } }; } };
  // On the list: the admin tools, no password, and the change is logged.
  w.user.email = 'lead@easdpa.org';
  let r = call({ action: 'query', table: 'reflections', filters: {} });
  assert.ok(r.ok && r.data.length === 2, JSON.stringify(r));
  // A teacher at the school domain but not on the list: refused for the admin tools...
  w.user.email = 'teacher@easdpa.org';
  r = call({ action: 'query', table: 'reflections', filters: {} });
  assert.ok(!r.ok && /staff password/i.test(r.error), JSON.stringify(r));
  // ...unless they have the shared password while PASSWORD_SIGN_IN is ON.
  r = call({ action: 'query', table: 'reflections', filters: {}, adminPassword: 'staffpw' });
  assert.ok(r.ok, JSON.stringify(r));
  // Unticking Active takes it away on the next call; the owner always counts.
  tabs['Staff Access'].rows[1][3] = false;
  w.user.email = 'lead@easdpa.org';
  assert.ok(!call({ action: 'query', table: 'reflections', filters: {} }).ok);
  w.user.email = 'owner@easdpa.org';
  assert.ok(call({ action: 'query', table: 'reflections', filters: {} }).ok);
  // Wrong passwords: alerts at 5.
  w.user.email = 'teacher@easdpa.org';
  for (let i = 0; i < 5; i++) call({ action: 'query', table: 'reflections', filters: {}, adminPassword: 'nope' + i });
  assert.ok(mails.some(m => /5 wrong staff passwords/.test(m.subject)), JSON.stringify(mails.map(m => m.subject)));
  // PASSWORD_SIGN_IN = OFF: the password stops working and the message says why.
  mkSheet('Settings', [['Key', 'Value', 'Notes'], ['PASSWORD_SIGN_IN', 'OFF', '']]);
  delete w.CacheService.__c; // the wrong-password count runs out
  r = call({ action: 'query', table: 'reflections', filters: {}, adminPassword: 'staffpw' });
  assert.ok(!r.ok && /not on the Trail Journal Staff Access list/.test(r.error), JSON.stringify(r));
  assert.throws(() => w.tjStaffSignIn('staffpw'), /turned off/);
  // Staff Actions: gated changes are logged; reads are not.
  call({ action: 'update', table: 'panel_cases', match: { id: 'c1' }, patch: { status: 'x' } });
  w.user.email = 'owner@easdpa.org';
  call({ action: 'update', table: 'panel_cases', match: { id: 'c1' }, patch: { status: 'x' } });
  const acts = (tabs['Staff Actions'] || { rows: [] }).rows;
  assert.ok(acts.some(a => a[2] === 'owner@easdpa.org' && /update panel_cases/.test(a[4]) && a[5] === 'row c1'), JSON.stringify(acts));
  assert.ok(!acts.some(a => /query reflections/.test(a[4])), 'reads are not logged');
});

test('trailjournal.org staff gates: Sign in with my school account opens STAFF_EXEC_URL (+ ?page=) only when this deployment cannot see a staff account', () => {
  const w = world();
  const real = w.webSetting_;
  w.webSetting_ = k => k === 'STAFF_EXEC_URL' ? 'https:' + '//script.google.com/a/macros/easdpa.org/s/STAFF/exec' : real(k);
  w.user.email = '';
  assert.strictEqual(w.staffLinkFor_('admin'), 'https:' + '//script.google.com/a/macros/easdpa.org/s/STAFF/exec?page=admin');
  assert.strictEqual(w.staffLinkFor_('panel'), 'https:' + '//script.google.com/a/macros/easdpa.org/s/STAFF/exec?page=panel');
  w.user.email = 'teacher@easdpa.org';
  assert.strictEqual(w.staffLinkFor_('admin'), '', 'already on the staff link: no button');
  w.user.email = '';
  w.webSetting_ = k => k === 'STAFF_EXEC_URL' ? '' : real(k);
  assert.strictEqual(w.staffLinkFor_('admin'), '', 'no staff link set: password only');
  const bridge = fs.readFileSync(path.join(dir, 'TjBridge.html'), 'utf8');
  assert.ok(bridge.includes('window.open(TJ_BOOT.staffLink'));
});

console.log('\n' + passed + ' tests passed');
