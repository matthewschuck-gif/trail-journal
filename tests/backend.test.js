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
  const ctx = { console, JSON, Math, Date, Object, Array, String, Number };
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

console.log('\n' + passed + ' tests passed');
