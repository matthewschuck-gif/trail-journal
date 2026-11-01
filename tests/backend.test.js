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

console.log('\n' + passed + ' tests passed');
