// Offline checks that every journal reaches the Sheet and the office (10_Records.gs).
// Fake Google services, real .gs code, made-up students.
//   node tests/records.test.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const dir = path.join(__dirname, '..', 'apps-script');
let passed = 0;
function test(name, fn) { fn(); passed++; console.log('ok - ' + name); }

function makeSheet(rows) {
  const s = {
    rows,
    getLastRow() { return this.rows.length; },
    getLastColumn() { return this.rows.reduce((m, r) => Math.max(m, r.length), 0); },
    getDataRange() { return { getValues: () => s.rows.map(r => r.slice()) }; },
    getRange(r, c, nr, nc) {
      nr = nr || 1; nc = nc || 1;
      return {
        getValues: () => { const o = []; for (let i = 0; i < nr; i++) { const row = s.rows[r - 1 + i] || []; o.push(Array.from({ length: nc }, (_, j) => row[c - 1 + j] == null ? '' : row[c - 1 + j])); } return o; },
        setValue: v => { s.rows[r - 1] = s.rows[r - 1] || []; s.rows[r - 1][c - 1] = v; },
      };
    },
    appendRow(row) { this.rows.push(row.slice()); },
  };
  return s;
}

// A Sheet as it is today: every SCHEMA column up to student_id, none of the new ones.
function world(aiReply) {
  const sent = [], props = {}, triggers = [];
  const ctx = { console, JSON, Math, Date, Object, Array, String, Number,
    MailApp: { sendEmail: o => sent.push(o) },
    Utilities: { formatDate: d => d.toISOString().slice(0, 16), sleep() {}, newBlob: (data, type, name) => ({ data, type, name }), getUuid: () => 'u' + Math.random() },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
    ScriptApp: { getProjectTriggers: () => [], deleteTrigger() {}, newTrigger: h => ({ timeBased: () => ({ everyMinutes: () => ({ create: () => triggers.push(h) }) }) }) },
    SpreadsheetApp: { getUi: () => { throw new Error('no ui in tests'); }, getActiveSpreadsheet: () => ({ getSheetByName: () => null }) },
    Session: { getEffectiveUser: () => ({ getEmail: () => 'me@x' }), getActiveUser: () => ({ getEmail: () => '' }) },
    UrlFetchApp: { fetch: () => ({ getResponseCode: () => 200, getContentText: () => JSON.stringify({ content: [{ text: JSON.stringify(aiReply || {}) }] }) }) },
    Logger: { log() {} },
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readdirSync(dir).filter(f => f.endsWith('.gs')).sort().map(f => fs.readFileSync(path.join(dir, f), 'utf8')).join('\n') + ';this.SCHEMA = SCHEMA;', ctx);
  const headers = ctx.SCHEMA.reflections.slice(0, ctx.SCHEMA.reflections.indexOf('student_id') + 1);
  const sheet = makeSheet([headers]);
  ctx.getSheet_ = () => sheet;
  ctx.getProp_ = k => ({ ANTHROPIC_API_KEY: 'k', APP_TOKEN: 'tok' })[k];
  ctx.getOfficeNotifyEmail_ = () => 'riggs@x,montagna@x,schuck@x';
  ctx.sent = sent; ctx.props = props; ctx.sheet = sheet; ctx.triggers = triggers;
  ctx.rowOf = sid => { const h = sheet.rows[0]; const r = sheet.rows.find(x => x[h.indexOf('session_id')] === sid); const o = {}; h.forEach((k, j) => { o[k] = r ? r[j] : undefined; }); return o; };
  return ctx;
}

const ago = min => new Date(Date.now() - min * 60000).toISOString();
const journal = (sid, extra) => Object.assign({ session_id: sid, initial: 'Z', student_name: 'Zed Quill', grade: '8', squad: 'GOLD', team: 'M', location: 'Hallways', peak: 'Be Safe',
  part1_json: { whatHappened: 'pushed in line', emotionsDuring: ['Frustrated'] }, part2_json: { hurtPeople: ['a classmate'], selectedConsequences: [] }, part3_json: { trait: 'Hope', followThrough: {} } }, extra || {});

test('every save lands on one row per session; new columns are added; a submitted journal is never turned back into a partial', () => {
  const w = world();
  w.upsertReflection_(journal('s1', { status: 'in_progress' }));
  w.upsertReflection_(journal('s1', { status: 'in_progress', part1_json: { whatHappened: 'more written' } }));
  assert.strictEqual(w.sheet.rows.length, 2, 'one row');
  const h = w.sheet.rows[0];
  assert.ok(h.indexOf('office_emailed') > -1 && h.indexOf('last_saved_at') === h.indexOf('office_emailed') + 1, 'columns added in SCHEMA order');
  assert.ok(/more written/.test(w.rowOf('s1').part1_json));
  w.upsertReflection_(journal('s1', { status: 'submitted' }));
  assert.strictEqual(w.upsertReflection_(journal('s1', { status: 'incomplete' })).kept, 'submitted');
  assert.strictEqual(w.rowOf('s1').status, 'submitted');
});

test('the AI summary finishes the submission on the server (report, goal, Clash flag, office email), and the page\'s copy is skipped', () => {
  const w = world({ insightLevel: 'High', studentGoal: 'Wait my turn', summary: 'Zed pushed in line and made a plan.' });
  w.upsertReflection_(journal('s2', { status: 'submitted' }));
  w.callAiProxy_('summary', { sessionId: 's2', name: 'Zed Quill', part1: {}, part2: {}, part3: {} });
  const r = w.rowOf('s2');
  assert.strictEqual(r.insight_level, 'High');
  assert.strictEqual(r.goal, 'Wait my turn');
  assert.strictEqual(r.clash_flagged, true);
  assert.ok(/WHAT HAPPENED/.test(r.reflection_text) && /pushed in line/.test(r.reflection_text), 'full written report on the row');
  assert.ok(/submitted/.test(r.office_emailed));
  assert.strictEqual(w.sent.length, 1);
  assert.strictEqual(w.sent[0].to, 'riggs@x,montagna@x,schuck@x');
  assert.ok(/New submission -- Zed Quill/.test(w.sent[0].subject));
  // The page closes nothing here, so its own email arrives too: skipped.
  const page = w.sendOfficeRecordEmail_({ studentLabel: 'Zed Quill', location: 'Hallways', tier: 'full', summaryText: 'x' });
  assert.strictEqual(page.duplicate, true);
  assert.strictEqual(w.sent.length, 1);
});

test('a partial (Ran Out of Time) always emails, is marked, and the later finished journal still gets its own email', () => {
  const w = world({ insightLevel: 'Medium' });
  w.upsertReflection_(journal('s3', { status: 'incomplete', ran_out_of_time_by: 'Ms. Lark' }));
  w.sendOfficeRecordEmail_({ studentLabel: 'Zed Quill', location: 'Hallways', incomplete: true, summaryText: 'partial' });
  assert.ok(/incomplete/.test(w.rowOf('s3').office_emailed));
  w.upsertReflection_(journal('s3', { status: 'submitted' }));
  w.callAiProxy_('summary', { sessionId: 's3', part1: {}, part2: {}, part3: {} });
  assert.strictEqual(w.sent.length, 2);
  assert.ok(/INCOMPLETE/.test(w.sent[0].subject) && /New submission/.test(w.sent[1].subject));
});

test('safety net: a journal left in progress for an hour goes out as INCOMPLETE; a submitted one the page never finished gets its AI summary and email', () => {
  const w = world({ insightLevel: 'Low', studentGoal: 'Ask for help' });
  w.props.records_since = String(Date.now() - 24 * 3600 * 1000);
  w.sheet.rows[0] = w.SCHEMA.reflections.slice(); // columns already there
  const add = o => { w.upsertReflection_(o); };
  add(journal('s4', { status: 'in_progress', created_at: ago(120) }));
  w.sheet.rows[w.sheet.rows.length - 1][w.sheet.rows[0].indexOf('last_saved_at')] = ago(90);
  add(journal('s5', { status: 'in_progress', created_at: ago(30) }));          // still being written
  add(journal('s6', { status: 'submitted', created_at: ago(40), student_name: 'Ivy Moss' })); // page closed before the AI came back
  w.recordSafetyNet();
  const s4 = w.rowOf('s4'), s5 = w.rowOf('s5'), s6 = w.rowOf('s6');
  assert.strictEqual(s4.status, 'incomplete'); assert.strictEqual(s4.ran_out_of_time_by, 'Auto (left unfinished)');
  assert.ok(String(s4.reflection_text).length > 100 && /incomplete/.test(s4.office_emailed));
  assert.strictEqual(s5.status, 'in_progress'); assert.ok(!s5.office_emailed, 'left alone while being written');
  assert.strictEqual(s6.insight_level, 'Low'); assert.strictEqual(s6.goal, 'Ask for help'); assert.ok(/submitted/.test(s6.office_emailed));
  assert.strictEqual(w.sent.length, 2);
  w.recordSafetyNet();
  assert.strictEqual(w.sent.length, 2, 'nothing sent twice');
});

test('one-time push: one email to the office with every journal as a spreadsheet, the missed ones written out and filled in on the Sheet', () => {
  const w = world();
  w.sheet.rows[0] = w.SCHEMA.reflections.slice();
  w.upsertReflection_(journal('old1', { status: 'submitted', created_at: '2026-09-10T14:00:00Z', insight_level: 'High', reflection_text: 'already emailed report' }));
  w.upsertReflection_(journal('old2', { status: 'submitted', created_at: '2026-10-20T14:00:00Z', student_name: 'Ivy Moss', insight_level: 'Medium' })); // no report: never emailed
  w.upsertReflection_(journal('old3', { status: 'incomplete', created_at: '2026-10-21T14:00:00Z', student_name: 'Rex Fern' }));                       // partial, no report
  const r = w.emailAllRecords_(Date.parse('2026-08-31T00:00:00'));
  assert.deepStrictEqual([r.total, r.missed], [3, 2]);
  assert.strictEqual(w.sent.length, 1);
  const mail = w.sent[0];
  assert.ok(/All journal records \+ 2 that were missed/.test(mail.subject));
  const csv = mail.attachments[0];
  assert.ok(/\.csv$/.test(csv.name));
  ['Zed Quill', 'Ivy Moss', 'Rex Fern', 'Was Never Emailed', 'Full Report'].forEach(s => assert.ok(csv.data.includes(s), s));
  assert.ok(/Ivy Moss/.test(mail.htmlBody) && /Rex Fern/.test(mail.htmlBody) && /INCOMPLETE/.test(mail.htmlBody));
  assert.ok(String(w.rowOf('old2').reflection_text).length > 100, 'filled in on the Sheet');
  assert.strictEqual(w.rowOf('old1').reflection_text, 'already emailed report', 'existing report left as it was');
  assert.strictEqual(w.rowOf('old1').office_emailed, 'before tracking');
  assert.ok(/resent once/.test(w.rowOf('old2').office_emailed));
  // The safety net (tracking from now) doesn't send the old ones again.
  w.props.records_since = String(Date.now());
  w.recordSafetyNet();
  assert.strictEqual(w.sent.length, 1);
});

test('the journal page can save through the router with no staff password (same as its insert/update today)', () => {
  const w = world();
  let out;
  w.jsonOut_ = o => { out = o; return o; };
  w.handleRequest_({ postData: { contents: JSON.stringify({ token: 'tok', action: 'upsertReflection', row: journal('s9', { status: 'in_progress' }) }) } }, 'POST');
  assert.ok(out && out.ok, JSON.stringify(out));
  assert.strictEqual(w.rowOf('s9').status, 'in_progress');
});

console.log('\n' + passed + ' tests passed');
