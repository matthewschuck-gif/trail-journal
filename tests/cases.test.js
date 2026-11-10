// Offline checks for Trailback Panel cases (12_TrailbackCases.gs): the status steps, the Master Log
// on the Accountability Sheet, the emails, the referral form and the calls from other tools.
// Fake Google services, real .gs code, made-up students.
//   node tests/cases.test.js
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
    Utilities: { formatDate: d => d.toISOString().slice(0, 10), getUuid: () => Math.random().toString(16).slice(2, 10) + '-3456-7890-abcd-ef1234567890' },
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
const offeredCase = { student_id: '801-2345', student_name: 'Pat Quill', grade: '8', squad: 'Gold', behavior_type: 'Disrespect',
  consequence_type: 'After School Detention', consequence_original_date: '2026-12-01', master_log_row: 'EMS-20261201-0042',
  referring_staff: 'Ms. AP', incident_summary: 'Argued with a sub and left class.', status: 'offered', source: 'admin_conference' };

// ── other tools ──
test('other tools: createTrailbackCase needs the shared key, works with the old token route OFF, and makes an "offered" case', () => {
  const w = world({ settings: [['LEGACY_API', 'OFF', '']] });
  w.props.TRAILBACK_API_KEY = 'k123';
  assert.ok(/Wrong or missing Trailback key/.test(w.post({ action: 'createTrailbackCase', key: 'nope', payload: offeredCase }).error));
  assert.ok(/retired|token/i.test(w.post({ action: 'createTrailbackCase', payload: offeredCase }).error), 'no key: normal route, refused');
  const r = w.post({ action: 'createTrailbackCase', key: 'k123', caller: 'Accountability', payload: Object.assign({}, offeredCase, { status: 'closed_completed' }) });
  assert.strictEqual(r.ok, true, JSON.stringify(r));
  assert.strictEqual(r.data.status, 'offered');
  assert.strictEqual(r.data.student_id, '8012345', '800# digits only');
  assert.ok(r.data.offered_at);
});

test('the same student + Master Log row while a case is open gives back that case, not a second one', () => {
  const w = world();
  const a = w.call({ action: 'createTrailbackCase', payload: offeredCase }).data;
  const b = w.call({ action: 'createTrailbackCase', payload: offeredCase }).data;
  assert.strictEqual(a.case_id, b.case_id);
  assert.strictEqual(b.existing, true);
  assert.strictEqual(w.tabs.panel_cases.rows.length, 2);
});

test('case actions are staff-only', () => {
  const w = world({ email: '' });
  ['trailbackCases', 'createTrailbackCase', 'updateTrailbackCase', 'trailbackCasesByStudent'].forEach(a =>
    assert.ok(/staff password/i.test(w.call({ action: a, payload: {} }).error), a));
});

// ── the whole path ──
test('full path: offer, consent (Master Log Held), Trail Journal auto-links, schedule, approve at the panel, submit, check-in, close completed (Resolved + anonymous Clash request)', () => {
  const w = world();
  const c = w.call({ action: 'createTrailbackCase', payload: Object.assign({}, offeredCase, { status: 'suggested' }) }).data;
  assert.strictEqual(c.status, 'suggested');
  assert.strictEqual(w.step(c.case_id, 'offer').data.case.status, 'offered');

  const consent = w.step(c.case_id, 'consent').data;
  assert.strictEqual(consent.case.status, 'consented');
  assert.deepStrictEqual(J(consent.warnings), []);
  assert.strictEqual(w.masterLog.rows[1][15], 'Held (Trailback)');
  assert.ok(/Trailback consent/.test(w.masterLog.rows[1][16]));

  // The student finishes a Trail Journal after consent: the board links it.
  const h = w.SCHEMA.reflections;
  const ref = h.map(k => ({ session_id: 'tj-77', student_id: '8012345', status: 'submitted', created_at: new Date(Date.now() + 1000).toISOString() })[k] || '');
  w.tabs.reflections.rows.push(ref);
  const board = w.call({ action: 'trailbackCases', payload: {} }).data;
  assert.strictEqual(board.linked, 1);
  const linked = board.cases.find(x => x.case_id === c.case_id);
  assert.strictEqual(linked.status, 'tj_done');
  assert.strictEqual(linked.trail_journal_id, 'tj-77');

  const wed = wednesdayAfter(7);
  assert.strictEqual(w.step(c.case_id, 'schedule', { date: wed }).data.case.scheduled_date, wed);
  assert.ok(w.call({ action: 'trailbackCases', payload: {} }).data.cases.find(x => x.case_id === c.case_id).status === 'scheduled');

  // At the panel
  const record = { session_id: 's-1', case_id: c.case_id, panel_date: wed, student_name: 'Pat Quill', student_id: '8012345', clash_squad: 'Gold',
    recommendation_narrative: 'Apologize to the sub and help in the classroom.', followup_person: 'Mr. Dean', fu_criteria: 'Apology + 2 help sessions',
    followup_dates: { checkin_3: '2020-01-06', checkin_10: '2020-01-15', checkin_30: '2099-01-01' } };
  const appr = w.call({ action: 'approvePanelPlan', payload: { record } }).data;
  assert.strictEqual(appr.case_id, c.case_id);
  assert.strictEqual(w.call({ action: 'trailbackCase', payload: { case_id: c.case_id } }).data.status, 'heard_approved');
  assert.strictEqual(w.call({ action: 'trailbackCase', payload: { case_id: c.case_id } }).data.panel_session_id, 's-1');
  w.call({ action: 'submitPanelSession', payload: { record } });
  const after = w.call({ action: 'trailbackCase', payload: { case_id: c.case_id } }).data;
  assert.strictEqual(after.status, 'repair_in_progress');
  assert.strictEqual(after.checkin_3, '2020-01-06');
  assert.strictEqual(after.checkin_person, 'Mr. Dean');
  assert.strictEqual(after.completion_criteria, 'Apology + 2 help sessions');

  const ci = w.step(c.case_id, 'checkin', { n: 3, note: 'Apology done.' }).data.case;
  assert.strictEqual(ci.checkin_notes['3'].on_time, false, 'logged after its due date');
  assert.strictEqual(w.step(c.case_id, 'fixit', { note: 'met with teacher' }).data.case.status, 'repair_in_progress');

  const done = w.step(c.case_id, 'close_completed').data;
  assert.strictEqual(done.case.status, 'closed_completed');
  assert.strictEqual(w.masterLog.rows[1][15], 'Resolved (Trailback)');
  assert.ok(done.case.clash_request_id);
  const q = w.call({ action: 'clashRequests' }).data;
  assert.strictEqual(q.length, 1);
  assert.strictEqual(q[0].squad, 'GOLD');
  assert.ok(!/Pat|Quill|P\.Q\.|8012345/.test(JSON.stringify(q)));
  assert.strictEqual(w.clashEvents.rows.length, 1, 'no Clash row until staff approve');
  assert.ok(done.case.history.length >= 9, 'every step is in the history');
});

test('steps out of order are refused', () => {
  const w = world();
  const c = w.call({ action: 'createTrailbackCase', payload: offeredCase }).data;
  assert.ok(/Can't schedule/.test(w.step(c.case_id, 'schedule', { date: wednesdayAfter(7) }).error));
  assert.ok(/Can't close/.test(w.step(c.case_id, 'close_completed').error));
  assert.ok(/Can't log a check-in/.test(w.step(c.case_id, 'checkin', { n: 3 }).error));
  w.step(c.case_id, 'consent');
  const thu = wednesdayAfter(7).replace(/\d\d$/, d => String(Number(d) + 1).padStart(2, '0'));
  assert.ok(/Wednesdays/.test(w.step(c.case_id, 'schedule', { date: thu }).error));
  assert.ok(/already passed/.test(w.step(c.case_id, 'schedule', { date: '2020-01-01' }).error));
});

test('schedule refuses a Wednesday on the No School Dates tab', () => {
  const wed = wednesdayAfter(14);
  const w = world();
  w.tabs['No School Dates'] = makeSheet([['Date', 'Note'], [wed, 'In-service']]);
  const c = w.call({ action: 'createTrailbackCase', payload: offeredCase }).data;
  w.step(c.case_id, 'consent');
  assert.ok(/No School Dates/.test(w.step(c.case_id, 'schedule', { date: wed }).error));
});

test('close reverted: needs a reason, Master Log back to Scheduled with a note, flags admin, never books a date; email only when ON', () => {
  for (const on of [false, true]) {
    const w = world({ settings: [['TRAILBACK_EMAILS', on ? 'ON' : 'OFF', ''], ['TRAILBACK_COORDINATOR_EMAIL', 'kelly@easdpa.org', '']] });
    const c = w.call({ action: 'createTrailbackCase', payload: offeredCase }).data;
    w.step(c.case_id, 'consent');
    w.call({ action: 'approvePanelPlan', payload: { record: { session_id: 's-r' + on, case_id: c.case_id, panel_date: '2026-12-09' } } });
    assert.ok(/why/.test(w.step(c.case_id, 'close_reverted').error));
    const r = w.step(c.case_id, 'close_reverted', { reason: 'Did not finish the apology' }).data;
    assert.strictEqual(r.case.status, 'closed_reverted');
    assert.ok(/new detention date/.test(r.case.admin_flag));
    assert.strictEqual(w.masterLog.rows[1][15], 'Scheduled');
    assert.ok(/reverted .*pick a new date/.test(w.masterLog.rows[1][16]));
    assert.strictEqual(String(w.masterLog.rows[1][2] || ''), '', 'no date booked');
    assert.strictEqual(w.sent.length, on ? 1 : 0);
    if (!on) assert.ok(r.warnings.some(x => /OFF/.test(x)));
    else assert.strictEqual(w.sent[0].to, 'kelly@easdpa.org');
    assert.strictEqual(w.step(c.case_id, 'clear_flag').data.case.admin_flag, '');
  }
});

test('declined before the panel; after consent the Master Log goes back to Scheduled; after the panel it is refused', () => {
  const w = world();
  const a = w.call({ action: 'createTrailbackCase', payload: offeredCase }).data;
  w.step(a.case_id, 'consent');
  const d = w.step(a.case_id, 'decline', { reason: 'Family said no' }).data;
  assert.strictEqual(d.case.status, 'declined');
  assert.strictEqual(w.masterLog.rows[1][15], 'Scheduled');
  const b = w.call({ action: 'createTrailbackCase', payload: Object.assign({}, offeredCase, { student_id: '8099999', master_log_row: 'EMS-20261201-0043' }) }).data;
  w.call({ action: 'approvePanelPlan', payload: { record: { session_id: 's-d', case_id: b.case_id, panel_date: '2026-12-09' } } });
  assert.ok(/Can't decline/.test(w.step(b.case_id, 'decline').error));
});

test('Master Log: a row for a different 800# is left alone, with a warning; a missing Entry ID is a warning, not an error', () => {
  const w = world();
  const c = w.call({ action: 'createTrailbackCase', payload: Object.assign({}, offeredCase, { master_log_row: 'EMS-20261201-0043' }) }).data;
  const r = w.step(c.case_id, 'consent').data;
  assert.strictEqual(r.case.status, 'consented');
  assert.ok(/different 800#/.test(r.warnings[0]));
  assert.strictEqual(w.masterLog.rows[2][15], 'Scheduled');
  const m = w.call({ action: 'createTrailbackCase', payload: Object.assign({}, offeredCase, { student_id: '8055555', master_log_row: 'EMS-NOPE' }) }).data;
  assert.ok(/No Master Log row/.test(w.step(m.case_id, 'consent').data.warnings[0]));
});

test('a panel started without a queued case gets one at approval', () => {
  const w = world();
  const appr = w.call({ action: 'approvePanelPlan', payload: { record: { session_id: 'walk-in', student_name: 'Lee Ray', student_id: '8077777', panel_date: '2026-12-09', clash_squad: 'Grey' } } }).data;
  assert.ok(appr.case_id);
  const c = w.call({ action: 'trailbackCase', payload: { case_id: appr.case_id } }).data;
  assert.strictEqual(c.status, 'heard_approved');
  assert.strictEqual(c.student_id, '8077777');
  assert.strictEqual(c.panel_session_id, 'walk-in');
  // The page can't move a session to another case afterward.
  w.call({ action: 'submitPanelSession', payload: { record: { session_id: 'walk-in', case_id: 'TB-OTHER', followup_dates: { checkin_3: '2099-01-01' } } } });
  assert.strictEqual(w.genericQuery_('panel_sessions', { session_id: 'walk-in' })[0].case_id, appr.case_id);
  assert.strictEqual(w.call({ action: 'trailbackCase', payload: { case_id: appr.case_id } }).data.status, 'repair_in_progress');
});

test('approving for a closed case is refused before anything is saved', () => {
  const w = world();
  const c = w.call({ action: 'createTrailbackCase', payload: offeredCase }).data;
  w.step(c.case_id, 'decline', { reason: 'no' });
  assert.ok(/Declined/.test(w.call({ action: 'approvePanelPlan', payload: { record: { session_id: 's-x', case_id: c.case_id } } }).error));
  assert.strictEqual(w.genericQuery_('panel_sessions', { session_id: 's-x' }).length, 0);
});

// ── emails ──
test('Tuesday brief: Wednesday cases, adults on Panel Members only, conflict line to Kelly; nothing sent while OFF; preview sends nothing', () => {
  const wed = wednesdayAfter(0);
  const w = world({ settings: [['TRAILBACK_EMAILS', 'OFF', '']] });
  w.tabs['Panel Members'] = makeSheet([['Name', 'Email', 'Role', 'Active'], ['Ms. AP', 'ap@easdpa.org', 'AP', true], ['Riley', 'riley@gomounts.org', 'Student', true],
    ['Mr. Dean', 'dean@easdpa.org', 'Dean', true], ['Old', 'old@easdpa.org', 'Teacher', false]]);
  const c = w.call({ action: 'createTrailbackCase', payload: Object.assign({}, offeredCase, { behavior_pattern: 'Third time this month' }) }).data;
  w.step(c.case_id, 'consent');
  // scheduled for the coming Wednesday (written directly: the schedule step refuses today if it is Wednesday)
  w.genericUpdate_('panel_cases', { case_id: c.case_id }, { status: 'scheduled', scheduled_date: wed });
  const off = J(w.trailbackTuesdayBrief_(false));
  assert.strictEqual(off.sent, false);
  assert.ok(/OFF/.test(off.why));
  assert.strictEqual(off.to, 'ap@easdpa.org,dean@easdpa.org', 'no student panelists, no inactive members');
  assert.ok(/Pat Quill/.test(off.html) && /Third time this month/.test(off.html) && /Argued with a sub/.test(off.html));
  assert.ok(/tell Kelly by <strong>Wednesday morning, before the panel/.test(off.html));
  assert.ok(/border-left:6px solid #8a5a00/.test(off.html) && /Gold squad/.test(off.html), 'the case card wears the Gold squad color');
  assert.strictEqual(w.sent.length, 0);
  w.tabs.Settings.rows.push(['TRAILBACK_EMAILS', 'ON', '']);
  w.tabs.Settings.rows.splice(1, 1);
  assert.strictEqual(J(w.trailbackTuesdayBrief_(true)).sent, false, 'preview never sends');
  assert.strictEqual(J(w.trailbackTuesdayBrief_(false)).sent, true);
  assert.strictEqual(w.sent.length, 1);
});

test('overdue check-ins: only past-due ones not logged yet; goes to the coordinator', () => {
  const w = world({ settings: [['TRAILBACK_EMAILS', 'ON', ''], ['TRAILBACK_COORDINATOR_EMAIL', 'kelly@easdpa.org', '']] });
  const c = w.call({ action: 'createTrailbackCase', payload: offeredCase }).data;
  w.genericUpdate_('panel_cases', { case_id: c.case_id }, { status: 'repair_in_progress', checkin_3: '2020-01-06', checkin_10: '2020-01-15', checkin_30: '2099-01-01', checkin_notes: JSON.stringify({ 3: { on_time: true } }) });
  const list = J(w.overdueCheckins_());
  assert.deepStrictEqual(list.map(o => o.n), [10]);
  const r = J(w.trailbackOverdueEmail_(false));
  assert.strictEqual(r.sent, true);
  assert.strictEqual(w.sent[0].to, 'kelly@easdpa.org');
  assert.ok(/10-school-day check-in was due 2020-01-15/.test(w.sent[0].htmlBody));
});

test('turning the timers on makes the Tuesday brief and the daily check', () => {
  const w = world();
  w.setupTrailbackTriggers();
  assert.deepStrictEqual(J(w.triggers), ['trailbackTuesdayBrief', 'trailbackDailyCheck']);
});

// ── referral form ──
function fakeForm(titles, collectsEmail) {
  const items = titles.map(t => ({ getTitle: () => t }));
  const add = type => () => { const it = { type, getTitle: () => it.title };
    it.setTitle = t => { it.title = t; return it; }; it.setRequired = r => { it.required = r; return it; };
    it.setHelpText = () => it; it.setChoiceValues = c => { it.choices = c; return it; }; items.push(it); return it; };
  return { items, getItems: () => items.slice(), collectsEmail: () => !!collectsEmail,
    addTextItem: add('text'), addParagraphTextItem: add('paragraph'), addListItem: add('list'), addMultipleChoiceItem: add('choice') };
}

test('referral form: adds only the missing questions, keeps ones already there (any wording), safe to run twice', () => {
  const w = world();
  const form = fakeForm(['Student Name', 'Student 800 Number', 'What have you already tried with this student?']);
  w.FormApp.openById = () => form;
  const r = J(w.addReferralQuestions());
  assert.deepStrictEqual(r.added, ['Behavior type', 'How many times, and on what dates?', 'Who was affected, and how?',
    'Will you share the impact at the panel?', 'Best time to reach you', 'Your name']);
  assert.strictEqual(r.found.length, 3);
  const share = form.items.find(i => i.title === 'Will you share the impact at the panel?');
  assert.deepStrictEqual(J(share.choices), ['In person', 'In writing', 'No']);
  assert.strictEqual(form.items.find(i => i.title === 'Behavior type').type, 'list');
  assert.deepStrictEqual(J(w.addReferralQuestions()).added, []);
  // every title it adds is read back by the parser
  const answers = {}; form.items.forEach(i => { answers[i.getTitle()] = 'x-' + i.getTitle(); });
  const c = J(w.referralToCase_(answers, ''));
  assert.ok(c.student_name && c.student_id && c.behavior_type && c.referring_staff && c.incident_summary && /Times and dates/.test(c.behavior_pattern) && /Tried/.test(c.behavior_pattern));
  assert.ok(c.referral_json.teacher_shares_impact && c.referral_json.best_time_to_reach);
});

test('referral form: no "Your name" question when the form collects email addresses', () => {
  const w = world();
  const form = fakeForm([], true);
  w.FormApp.openById = () => form;
  assert.ok(!J(w.addReferralQuestions()).added.includes('Your name'));
});

test('teacher referral form: one submission becomes a "suggested" teacher_referral case', () => {
  const w = world();
  const answers = {
    'Student name': 'Sam Ortiz', 'Student 800#': '8044444', 'Behavior type': 'Disruption',
    'How many times, and on what dates?': '3 times: 11/2, 11/5, 11/9', 'What have you tried?': 'Seat change, call home',
    'Who was affected, and how?': 'Two tablemates could not finish their lab', 'Will you share the impact at the panel?': 'In writing',
    'Best time to reach you': '3rd period', 'Your name': 'Mr. Teacher',
  };
  const item = (t, v) => ({ getItem: () => ({ getTitle: () => t }), getResponse: () => v });
  w.onTrailbackReferral({ response: { getItemResponses: () => Object.keys(answers).map(k => item(k, answers[k])), getRespondentEmail: () => 'teacher@easdpa.org' } });
  const c = w.call({ action: 'trailbackCases', payload: {} }).data.cases[0];
  assert.strictEqual(c.status, 'suggested');
  assert.strictEqual(c.source, 'teacher_referral');
  assert.strictEqual(c.student_name, 'Sam Ortiz');
  assert.strictEqual(c.student_id, '8044444');
  assert.strictEqual(c.behavior_type, 'Disruption');
  assert.strictEqual(c.referring_staff, 'Mr. Teacher');
  assert.ok(/tablemates/.test(c.incident_summary));
  assert.ok(/11\/2/.test(c.behavior_pattern) && /Seat change/.test(c.behavior_pattern));
  assert.strictEqual(c.referral_json.teacher_shares_impact, 'In writing');
  assert.strictEqual(c.referral_json.best_time_to_reach, '3rd period');
});

// ── Climb-er ──
test('trailbackCasesByStudent: status, dates and outcome only (no names, notes, plan or reasons)', () => {
  const w = world();
  w.props.TRAILBACK_API_KEY = 'k1';
  const c = w.call({ action: 'createTrailbackCase', payload: Object.assign({}, offeredCase, { completion_criteria: 'secret plan' }) }).data;
  w.step(c.case_id, 'decline', { reason: 'Family said no, private reason' });
  const r = w.post({ action: 'trailbackCasesByStudent', key: 'k1', payload: { student_id: '8012345' } });
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.data[0].outcome, 'declined');
  const text = JSON.stringify(r.data);
  assert.ok(!/Pat|Quill|private|history|Argued|secret|reason/i.test(text), text);
  assert.ok(/student_id \(800#\) is required/.test(w.post({ action: 'trailbackCasesByStudent', key: 'k1', payload: {} }).error));
});

// ── board ──
test('board: weekly counts and the coming Wednesday docket', () => {
  const w = world();
  const c = w.call({ action: 'createTrailbackCase', payload: offeredCase }).data;
  w.step(c.case_id, 'consent');
  const board = w.call({ action: 'trailbackCases', payload: {} }).data;
  assert.strictEqual(board.counts.offered, 1);
  assert.strictEqual(board.counts.accepted, 1);
  assert.strictEqual(new Date(board.wednesday + 'T12:00:00Z').getUTCDay(), 3);
  assert.ok(Array.isArray(board.docket));
});

test('test cases: add and remove only touch TB-TEST rows', () => {
  const w = world();
  w.call({ action: 'createTrailbackCase', payload: offeredCase });
  w.addTrailbackTestCase();
  assert.strictEqual(w.tabs.panel_cases.rows.length, 3);
  w.removeTrailbackTestCases();
  assert.strictEqual(w.tabs.panel_cases.rows.length, 2);
  assert.ok(!/TB-TEST/.test(JSON.stringify(w.tabs.panel_cases.rows)));
});

console.log('\n' + passed + ' tests passed');
