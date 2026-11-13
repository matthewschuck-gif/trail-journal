// Offline checks for the Trailback repair plan (14_RepairPlan.gs): questions, the plan on the case, check-ins, reminders.
// Fake Google services, real .gs code, made-up students.
//   node tests/repair.test.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const dir = path.join(__dirname, '..', 'apps-script');
let passed = 0;
function test(name, fn) { fn(); passed++; console.log('ok - ' + name); }
const J = x => JSON.parse(JSON.stringify(x));

function makeSheet(rows) {
  const s = {
    rows,
    getLastRow() { return this.rows.length; },
    getLastColumn() { return this.rows.reduce((m, r) => Math.max(m, r.length), 0); },
    getDataRange() { return { getValues: () => s.rows.map(r => r.slice()), getDisplayValues: () => s.rows.map(r => r.map(v => v == null ? '' : String(v))) }; },
    getRange(r, c, nr, nc) {
      nr = nr || 1; nc = nc || 1;
      const range = {
        getValues: () => { const o = []; for (let i = 0; i < nr; i++) { const row = s.rows[r - 1 + i] || []; o.push(Array.from({ length: nc }, (_, j) => row[c - 1 + j] == null ? '' : row[c - 1 + j])); } return o; },
        setValue: v => { s.rows[r - 1] = s.rows[r - 1] || []; s.rows[r - 1][c - 1] = v; return range; },
        setValues: vals => { vals.forEach((row, i) => { s.rows[r - 1 + i] = s.rows[r - 1 + i] || []; row.forEach((v, j) => { s.rows[r - 1 + i][c - 1 + j] = v; }); }); return range; },
        setFontWeight: () => range, setBackground: () => range, setFontColor: () => range,
      };
      return range;
    },
    appendRow(row) { this.rows.push(row.slice()); },
    deleteRow(i) { this.rows.splice(i - 1, 1); },
    setFrozenRows() {},
  };
  return s;
}

const ACC_ID = '1IVFnkQQvSvNSY_OZPMutPzkTqmJ1j5XYRoT9YvNy-v4';
function mlRow(entryId, sid, status) { const r = new Array(34).fill(''); r[0] = entryId; r[6] = 'Quill, Pat'; r[15] = status || 'Scheduled'; r[33] = sid || ''; return r; }

function world(opts) {
  opts = opts || {};
  const tabs = {}, sent = [], props = {}, triggers = [];
  const user = { email: opts.email === undefined ? 'kelly@easdpa.org' : opts.email };
  const ss = { getSheetByName: n => tabs[n] || null, insertSheet: n => (tabs[n] = makeSheet([])) };
  const masterLog = makeSheet([new Array(34).fill('h'), mlRow('EMS-20261201-0042', '8012345'), mlRow('EMS-20261201-0043', '8099999')]);
  const clashEvents = makeSheet([['id', 'name', 'type', 'date', 'note', 'multiplier', 'BLACK', 'GOLD', 'GREY', 'PURPLE']]);
  const ctx = { console, JSON, Math, Date, Object, Array, String, Number,
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      openById: id => id === ACC_ID ? { getSheetByName: n => n === 'Master Log' ? masterLog : null } : { getSheetByName: n => n === 'events' ? clashEvents : null },
      getUi: () => { throw new Error('no ui'); },
    },
    Session: { getActiveUser: () => ({ getEmail: () => user.email }), getEffectiveUser: () => ({ getEmail: () => 'me@easdpa.org' }) },
    Utilities: { formatDate: d => d.toISOString().slice(0, 10), getUuid: () => Math.random().toString(16).slice(2, 10) + '-3456-7890-abcd-ef1234567890', newBlob: (text, type, name) => ({ text, type, name }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    MailApp: { sendEmail: o => sent.push(o) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
    ScriptApp: { getProjectTriggers: () => [], deleteTrigger() {}, WeekDay: { TUESDAY: 'TUE' },
      newTrigger: h => { const t = { timeBased: () => t, onWeekDay: () => t, atHour: () => t, nearMinute: () => t, everyDays: () => t, forForm: () => t, onFormSubmit: () => t, create: () => triggers.push(h) }; return t; } },
    FormApp: { openById: () => ({}) },
    HtmlService: { createHtmlOutput: h => ({ setWidth: () => ({ setHeight: () => h }) }) },
    Logger: { log() {} },
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readdirSync(dir).filter(f => f.endsWith('.gs')).sort().map(f => fs.readFileSync(path.join(dir, f), 'utf8')).join('\n') + ';this.SCHEMA = SCHEMA;', ctx);
  tabs.Settings = makeSheet([['Key', 'Value', 'Notes']].concat(opts.settings || []));
  tabs.reflections = makeSheet([ctx.SCHEMA.reflections.slice()]);
  ctx.getProp_ = k => ({ APP_TOKEN: 'tok', ADMIN_PASSWORD: 'staffpw' })[k];
  ctx.call = body => J(JSON.parse(ctx.tjCall(JSON.stringify(body))));
  ctx.post = body => { let out; ctx.jsonOut_ = o => { out = o; return o; }; ctx.handleRequest_({ postData: { contents: JSON.stringify(body) } }, 'POST'); return J(out); };
  ctx.step = (case_id, op, extra) => ctx.call({ action: 'updateTrailbackCase', payload: Object.assign({ case_id, op }, extra || {}) });
  Object.assign(ctx, { tabs, sent, props, triggers, user, masterLog, clashEvents });
  return ctx;
}

// A Wednesday a few weeks out (never in the past, whatever day the tests run).
function wednesdayAfter(days) {
  const d = new Date(); d.setDate(d.getDate() + days);
  while (d.getDay() !== 3) d.setDate(d.getDate() + 1);
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

const aCase = { student_id: '8012345', student_name: 'Pat Quill', grade: '8', squad: 'Gold', behavior_type: 'Technology',
  consequence_type: 'After School Detention', consequence_original_date: '2026-12-01', master_log_row: 'EMS-20261201-0042',
  referring_staff: 'Ms. AP', incident_summary: 'Posted a photo of a classmate.', status: 'offered', source: 'admin_conference' };

// A case through the panel with a two-action plan. Returns the case id.
function heardWithPlan(w, plan, extra) {
  const c = w.call({ action: 'createTrailbackCase', payload: aCase }).data;
  w.step(c.case_id, 'consent');
  const wed = wednesdayAfter(7);
  w.step(c.case_id, 'schedule', { date: wed });
  const record = Object.assign({ session_id: 's-' + c.case_id, case_id: c.case_id, panel_date: wed, student_name: 'Pat Quill', student_id: '8012345', clash_squad: 'Gold',
    recommendation_narrative: 'Take it down and teach others.', theme_sentence: 'Posts first, thinks later.', need_bucket: 'Awareness',
    followup_dates: { checkin_3: '2020-01-06', checkin_10: '2020-01-15', checkin_30: '2099-01-01' }, plan_actions: plan }, extra || {});
  assert.ok(w.call({ action: 'approvePanelPlan', payload: { record } }).ok);
  const r = w.call({ action: 'submitPanelSession', payload: { record } });
  assert.ok(r.ok, r.error);
  return c.case_id;
}
const PLAN = [
  { menu: 'Repair', action: 'Delete the post and apologize to the classmate', test: 'The classmate says it happened', due: '2099-01-10', adult: 'Mr. Dean <dean@easdpa.org>', rrr: { related: true, respectful: true, reasonable: true } },
  { menu: 'Reflect', action: 'PSA slideshow about posting photos', test: 'Shown in advisory', due: '2099-01-12', adult: 'coach@easdpa.org', rrr: { related: true, respectful: true, reasonable: true } },
];

test('the Panel Questions tab: the v2 sheet for every case, plus the case\'s behavior type, grouped by stage; staff only', () => {
  const w = world();
  const q = w.call({ action: 'panelQuestions', payload: { behavior: 'Technology' } });
  assert.ok(q.ok, q.error);
  const st = q.data.stages;
  assert.ok(st.Strengths.some(x => /What are you good at/.test(x.q) && x.when === 'Ask first'));
  assert.ok(st.Impact.some(x => /Who saw it online/.test(x.q)), 'the Technology starters are in');
  assert.ok(!st.Impact.some(x => /in their body/.test(x.q)), 'other behavior types are not');
  assert.ok(st.Impact.some(x => /stood out in what you just heard/.test(x.q) && /after/i.test(x.when)));
  assert.strictEqual(st['Back pocket'].length, 6);
  assert.ok(w.tabs['Panel Questions'].rows.length > 50, 'seeded on the tab');
  w.tabs['Panel Questions'].rows.forEach(r => { if (/What are you good at/.test(r[4])) r[0] = false; });
  assert.ok(!w.call({ action: 'panelQuestions', payload: {} }).data.stages.Strengths.some(x => /What are you good at/.test(x.q)), 'untick Show to hide one');
  w.user.email = '';
  assert.ok(/staff/i.test(w.call({ action: 'panelQuestions', payload: {} }).error));
});

test('the plan goes onto the case at submit: one entry per action (max 3), each tracked; the check-in person defaults to the first adult', () => {
  const w = world();
  const id = heardWithPlan(w, PLAN.concat([PLAN[0], PLAN[1]]));
  const c = w.call({ action: 'trailbackCase', payload: { case_id: id } }).data;
  const a = JSON.parse(c.repair_actions);
  assert.strictEqual(a.length, 3, 'one to three actions');
  assert.deepStrictEqual(a.map(x => x.id), ['a1', 'a2', 'a3']);
  assert.ok(a.every(x => x.status === 'open'));
  assert.strictEqual(a[0].test, 'The classmate says it happened');
  assert.strictEqual(c.checkin_person, 'Mr. Dean <dean@easdpa.org>');
  assert.ok(/apologize to the classmate: The classmate says it happened/.test(c.completion_criteria), 'the tests become the completion criteria');
  const s = w.tabs.panel_sessions.rows; const h = s[0];
  assert.strictEqual(s[s.length - 1][h.indexOf('need_bucket')], 'Awareness');
  assert.ok(/PSA slideshow/.test(s[s.length - 1][h.indexOf('plan_actions')]));
});

test('one-tap check-in: only the case\'s adults (or leads) can; it logs the due check-in; Needs help tells the coordinator; every action done closes the case', () => {
  const w = world({ settings: [['TRAILBACK_EMAILS', 'ON', ''], ['TRAILBACK_COORDINATOR_EMAIL', 'kelly@easdpa.org', '']] });
  const id = heardWithPlan(w, PLAN);
  w.user.email = 'stranger@easdpa.org';
  assert.strictEqual(w.call({ action: 'repairCheckins', payload: {} }).data.cases.length, 0);
  assert.ok(/not a check-in adult/.test(w.call({ action: 'repairCheckin', payload: { case_id: id, overall: 'on_track' } }).error));
  w.user.email = 'dean@easdpa.org';
  const list = w.call({ action: 'repairCheckins', payload: {} }).data;
  assert.strictEqual(list.cases.length, 1);
  assert.strictEqual(list.cases[0].next_checkin.n, 3, 'the 3-day check-in is due');
  w.sent.length = 0;
  const r1 = w.call({ action: 'repairCheckin', payload: { case_id: id, actions: { a1: 'done', a2: 'partly' }, notes: { a2: 'half the slides' }, overall: 'needs_help', note: 'Needs a ride to stay after.' } });
  assert.ok(r1.ok, r1.error);
  assert.strictEqual(r1.data.logged, 3);
  assert.strictEqual(r1.data.closed, false);
  assert.ok(w.sent.some(m => m.to === 'kelly@easdpa.org' && /needs help/.test(m.subject)), 'the coordinator hears about Needs help');
  let c = w.call({ action: 'trailbackCase', payload: { case_id: id } }).data;
  assert.ok(/Needs help/.test(c.checkin_notes['3'].note) && /PSA slideshow about posting photos: partly/.test(c.checkin_notes['3'].note));
  assert.deepStrictEqual(JSON.parse(c.repair_actions).map(a => a.status), ['done', 'partly']);
  // The second adult finishes it: every action done, the case closes as completed by itself.
  w.user.email = 'coach@easdpa.org';
  const r2 = w.call({ action: 'repairCheckin', payload: { case_id: id, actions: { a2: 'done' }, overall: 'on_track' } });
  assert.ok(r2.ok, r2.error);
  assert.strictEqual(r2.data.logged, 10);
  assert.strictEqual(r2.data.closed, true);
  c = w.call({ action: 'trailbackCase', payload: { case_id: id } }).data;
  assert.strictEqual(c.status, 'closed_completed');
  assert.strictEqual(w.masterLog.rows[1][15], 'Resolved (Trailback)');
  assert.ok(c.clash_request_id, 'the anonymous Clash point request, like closing by hand');
  // With TRAILBACK_AUTO_CLOSE OFF, it stays open for staff to close.
  const w2 = world({ settings: [['TRAILBACK_AUTO_CLOSE', 'OFF', '']] });
  const id2 = heardWithPlan(w2, [PLAN[0]]);
  w2.user.email = 'dean@easdpa.org';
  assert.strictEqual(w2.call({ action: 'repairCheckin', payload: { case_id: id2, actions: { a1: 'done' }, overall: 'on_track' } }).data.closed, false);
});

test('reminders: the check-in link the morning it is due, a heads-up 2 school days before an action is due, a past-due note to the coordinator; each once', () => {
  const w = world({ settings: [['TRAILBACK_EMAILS', 'ON', ''], ['TRAILBACK_COORDINATOR_EMAIL', 'kelly@easdpa.org', '']] });
  const today = w.todayYmd_();
  const plus = n => { const d = new Date(today + 'T12:00:00'); let k = 0; while (k < n) { d.setDate(d.getDate() + 1); if (d.getDay() && d.getDay() < 6) k++; } return w.ymdOf_(d); };
  const id = heardWithPlan(w, [Object.assign({}, PLAN[0], { due: plus(2) }), Object.assign({}, PLAN[1], { due: '2020-01-02' })],
    { followup_dates: { checkin_3: today, checkin_10: '2099-01-15', checkin_30: '2099-02-01' } });
  w.sent.length = 0;
  w.repairReminders_(false);
  const subj = w.sent.map(m => m.subject + ' -> ' + m.to);
  assert.ok(subj.some(s => /check-in today/.test(s) && /dean@easdpa.org/.test(s) && /coach@easdpa.org/.test(s)), 'the adults get the check-in link');
  assert.ok(w.sent.find(m => /check-in today/.test(m.subject)).htmlBody.includes('page=checkin'), 'with the one-tap link');
  assert.ok(subj.some(s => /due in 2 school days/.test(s) && /dean@easdpa.org/.test(s) && /kelly@easdpa.org/.test(s)));
  assert.ok(subj.some(s => /past due/.test(s) && /kelly@easdpa.org/.test(s)));
  const n = w.sent.length;
  w.repairReminders_(false);
  assert.strictEqual(w.sent.length, n, 'each reminder once');
  assert.strictEqual(w.call({ action: 'trailbackCase', payload: { case_id: id } }).data.status, 'repair_in_progress', 'nothing reverts by itself');
});

console.log('\n' + passed + ' tests passed');
