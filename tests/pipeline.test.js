// Offline checks for the Trailback autopilot (17_Pipeline.gs): spot, ask, journal, schedule, seats, glimpse.
// Fake Google services, real .gs code, made-up students.
//   node tests/pipeline.test.js
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

const ML_HEADERS = ['Entry ID', 'Entry Date', 'Consequence Date', 'Consequence Type', 'Assigned By (name)', 'Assigned By (role)', 'Student Name', 'Grade', 'Squad', 'Homeroom',
  'Reason / Notes', 'Monitor / Location', 'Transport Type', 'Bus # (if activity bus)', 'Bus Stop (if activity bus)', 'Status', 'Status Notes', 'Block Slot # (of 6)',
  'Admin-Reserved Slot?', 'Follow-Up Required?', 'Follow-Up Role', 'Follow-Up Assignee', 'Follow-Up Due', 'Follow-Up Completed?', 'Follow-Up Notes', 'Teacher Checklist ID',
  'Reflection Given?', 'Reflection Type', 'Missed Work Assigned? (Teacher LD)', 'Repair Plan Completed? (Admin LD)', 'Reflection Returned?', 'Parent Contacted?',
  'Contact Date & Method', 'Student 800#', 'Behavior Type'];
const ago = n => { const d = new Date(); d.setDate(d.getDate() - n); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); };
function ml(id, sid, name, type, day, cons, status) {
  const r = ML_HEADERS.map(() => '');
  const set = (h, v) => { r[ML_HEADERS.indexOf(h)] = v; };
  set('Entry ID', id); set('Entry Date', day); set('Consequence Date', day); set('Consequence Type', cons || 'Teacher Conference / Warning'); set('Student Name', name);
  set('Grade', '8'); set('Squad', 'Gold'); set('Status', status || 'Scheduled'); set('Student 800#', sid); set('Behavior Type', type); set('Reason / Notes', 'Talking over the teacher');
  return r;
}
const settings = [['TRAILBACK_EMAILS', 'ON', ''], ['TRAILBACK_COORDINATOR_EMAIL', 'kelly@easdpa.org', ''], ['ACCOUNTABILITY_SHEET_ID', '1IVFnkQQvSvNSY_OZPMutPzkTqmJ1j5XYRoT9YvNy-v4', '']];

test('spot: 3 of the same behavior in 30 days suggests a case (skip words never do); a Trail Journal flag suggests one; never twice; one digest', () => {
  const w = world({ settings });
  w.masterLog.rows = [ML_HEADERS.slice(),
    ml('E1', '8012345', 'Quill, Pat', 'Disruption', ago(20)), ml('E2', '8012345', 'Quill, Pat', 'Disruption', ago(9)), ml('E3', '8012345', 'Quill, Pat', 'Disruption', ago(2), 'Admin Lunch Detention (LD)'),
    ml('E4', '8099999', 'Rowe, Lee', 'Bullying', ago(5)), ml('E5', '8099999', 'Rowe, Lee', 'Bullying', ago(4)), ml('E6', '8099999', 'Rowe, Lee', 'Bullying', ago(3)),
    ml('E7', '8055555', 'Moss, Kai', 'Disruption', ago(50)), ml('E8', '8055555', 'Moss, Kai', 'Disruption', ago(3)), ml('E9', '8055555', 'Moss, Kai', 'Disruption', ago(2), '', 'Cancelled')];
  const made = J(w.patternSuggestions_());
  assert.strictEqual(made.length, 1, 'only Pat: Lee is a skip word, Kai has one old and one cancelled');
  assert.strictEqual(made[0].student_name, 'Pat Quill');
  assert.strictEqual(made[0].source, 'pattern');
  assert.strictEqual(made[0].status, 'suggested');
  assert.strictEqual(made[0].master_log_row, 'E3', 'the detention is the one held');
  assert.ok(/3 Disruption records in 30 days/.test(made[0].behavior_pattern));
  assert.strictEqual(w.patternSuggestions_().length, 0, 'never twice');
  // A Trail Journal flag.
  const h = w.SCHEMA.reflections;
  w.tabs.reflections.rows.push(h.map(k => ({ session_id: 'tj-9', student_id: '8077777', student_name: 'Ash Fern', status: 'submitted', prior_count: 2, insight_level: 'Medium', created_at: new Date().toISOString() })[k] || ''));
  w.tabs.reflections.rows.push(h.map(k => ({ session_id: 'tj-10', student_id: '8066666', status: 'submitted', prior_count: 0, insight_level: 'Low', created_at: new Date().toISOString() })[k] || ''));
  const j = J(w.journalSuggestions_());
  assert.deepStrictEqual(j.map(c => c.student_id), ['8077777'], 'a first Low journal alone is not enough');
  assert.strictEqual(w.call({ action: 'trailbackCase', payload: { case_id: j[0].case_id } }).data.trail_journal_id, 'tj-9');
  w.sent.length = 0;
  w.suggestionDigest_(made.concat(j));
  assert.ok(w.sent.length === 1 && /2 new suggestions/.test(w.sent[0].subject) && w.sent[0].to === 'kelly@easdpa.org');
});

test('ask: offering sends the family a one-time consent link; Yes records consent and holds the detention, then the student gets the Trail Journal link', () => {
  const w = world({ settings });
  const c = w.call({ action: 'createTrailbackCase', payload: { student_id: '8012345', student_name: 'Pat Quill', master_log_row: 'EMS-20261201-0042', consequence_type: 'Admin Lunch Detention', family_email: 'Parent Quill <pq@example.org>' } }).data;
  w.sent.length = 0;
  const off = w.step(c.case_id, 'offer');
  assert.ok(off.ok, off.error);
  const mail = w.sent.find(m => m.to === 'pq@example.org');
  assert.ok(mail && /page=consent&(amp;)?t=/.test(mail.htmlBody), 'the family gets the link');
  const token = mail.htmlBody.match(/t=([a-z0-9]+)/)[1];
  w.user.email = ''; // the family isn't signed in
  const look = w.call({ action: 'consentLookup', payload: { t: token } }).data;
  assert.deepStrictEqual([look.first, look.open], ['Pat', true]);
  assert.ok(!JSON.stringify(look).includes('Quill'), 'first name only');
  w.sent.length = 0;
  const yes = w.call({ action: 'consentAnswer', payload: { t: token, answer: 'yes', name: 'Robin Quill' } });
  assert.ok(yes.ok, yes.error);
  w.user.email = 'kelly@easdpa.org';
  const after = w.call({ action: 'trailbackCase', payload: { case_id: c.case_id } }).data;
  assert.strictEqual(after.status, 'consented');
  assert.strictEqual(w.masterLog.rows[1][15], 'Held (Trailback)');
  assert.ok(/Family via link \(Robin Quill\)/.test(after.history[after.history.length - 1].by));
  assert.ok(w.sent.some(m => m.to === '8012345@gomounts.org' && /Trail Journal/.test(m.subject)), 'the student gets the Trail Journal link');
  w.user.email = '';
  assert.ok(/already been used/.test(w.call({ action: 'consentAnswer', payload: { t: token, answer: 'no' } }).error), 'one time only');
  // No family email: the coordinator is asked to call.
  w.user.email = 'kelly@easdpa.org';
  const c2 = w.call({ action: 'createTrailbackCase', payload: { student_id: '8099999', student_name: 'Lee Rowe', master_log_row: 'EMS-20261201-0043' } }).data;
  w.sent.length = 0;
  w.step(c2.case_id, 'offer');
  assert.ok(w.sent.some(m => m.to === 'kelly@easdpa.org' && /call the family/.test(m.subject)));
  // No: declined.
  w.step(c2.case_id, 'edit', { fields: { family_email: 'lr@example.org' } });
  w.sent.length = 0;
  w.step(c2.case_id, 'send_consent');
  const t2 = w.sent.find(m => m.to === 'lr@example.org').htmlBody.match(/t=([a-z0-9]+)/)[1];
  w.call({ action: 'consentAnswer', payload: { t: t2, answer: 'no' } });
  assert.strictEqual(w.call({ action: 'trailbackCase', payload: { case_id: c2.case_id } }).data.status, 'declined');
});

test('schedule + seats: Trail Journal done puts the case on the next open Wednesday; seat call and notices go out; a second seat call 2 school days before if short', () => {
  const w = world({ settings: settings.concat([['PANEL_MAX_CASES', '1', '']]) });
  w.tabs['Panel Members'] = makeSheet([['Name', 'Email', 'Role', 'Active'], ['Kelly', 'kelly@easdpa.org', 'Lead', 'TRUE'], ['Mr. Dean', 'dean@easdpa.org', 'Staff', 'TRUE']]);
  const mk = (sid, ml) => { const c = w.call({ action: 'createTrailbackCase', payload: { student_id: sid, student_name: 'Pat Quill', master_log_row: ml } }).data; w.step(c.case_id, 'consent'); w.step(c.case_id, 'link_tj', { trail_journal_id: 'tj-' + sid }); return c.case_id; };
  const a = mk('8012345', 'EMS-20261201-0042'), b = mk('8099999', 'EMS-20261201-0043');
  w.sent.length = 0;
  const done = J(w.autoSchedule_());
  assert.deepStrictEqual(done, [a, b]);
  const ca = w.call({ action: 'trailbackCase', payload: { case_id: a } }).data, cb = w.call({ action: 'trailbackCase', payload: { case_id: b } }).data;
  assert.strictEqual(new Date(ca.scheduled_date + 'T12:00:00').getDay(), 3, 'a Wednesday');
  assert.ok(w.schoolDaysBetween_(w.todayYmd_(), ca.scheduled_date) >= 2, 'at least 2 school days out');
  assert.ok(cb.scheduled_date > ca.scheduled_date, 'one case per Wednesday (PANEL_MAX_CASES 1): the next one goes a week later');
  assert.ok(w.sent.some(m => /Seats open/.test(m.subject) && /dean@easdpa.org/.test(m.to)), 'seat call to the Panel Members');
  assert.ok(!w.sent.find(m => /Seats open/.test(m.subject)).htmlBody.includes('Quill'), 'last initial only in the seat call');
  assert.ok(w.sent.some(m => m.to === '8012345@gomounts.org' && /Your Trailback Panel/.test(m.subject)), 'the student is told');
  // Two school days before, with nobody signed up: a second call, once.
  w.todayYmd_ = () => w.panelSchoolDayBefore_(w.panelSchoolDayBefore_(ca.scheduled_date));
  w.sent.length = 0;
  w.lowSeatCalls_(); w.lowSeatCalls_();
  assert.strictEqual(w.sent.filter(m => /Still need panelists/.test(m.subject)).length, 1);
});

test('the glimpse shows only to those signed up for the case', () => {
  const w = world({ settings });
  const c = w.call({ action: 'createTrailbackCase', payload: { student_id: '8012345', student_name: 'Pat Quill', master_log_row: 'EMS-20261201-0042', incident_summary: 'Talked over the sub.' } }).data;
  w.step(c.case_id, 'consent');
  w.step(c.case_id, 'schedule', { date: wednesdayAfter(7) });
  w.genericUpdate_('panel_cases', { case_id: c.case_id }, { climber_brief: JSON.stringify({ profile: { team: 'Summit' }, strengths: { counts: [{ label: 'Shout-outs', year: 4 }], activities: ['Robotics'] }, pattern: { counts: [{ label: 'Office referrals', year: 3, recent: 1 }] } }) });
  w.user.email = 'ann_one@easdpa.org';
  assert.ok(!w.call({ action: 'panelBoard', payload: {} }).data.cases[0].glimpse, 'not before signing up');
  w.call({ action: 'panelSignUp', payload: { caseId: c.case_id } });
  const g = w.call({ action: 'panelBoard', payload: {} }).data.cases[0].glimpse;
  assert.ok(g && g.incident === 'Talked over the sub.' && g.team === 'Summit');
  assert.ok(g.strengths.includes('Shout-outs: 4') && g.strengths.some(s => /Robotics/.test(s)));
  assert.ok(/Office referrals: 3 this year/.test(g.yearPattern[0]));
});

test('outcomes: before vs after the panel from the Master Log, the same days before and after, by behavior; leads and the coordinator only', () => {
  const w = world({ settings: settings.concat([['PANEL_ADMINS', 'lead@easdpa.org', '']]) });
  const c = w.call({ action: 'createTrailbackCase', payload: { student_id: '8012345', student_name: 'Pat Quill', behavior_type: 'Disruption', master_log_row: 'E3' } }).data;
  w.genericUpdate_('panel_cases', { case_id: c.case_id }, { status: 'closed_completed', scheduled_date: ago(100), closed_at: ago(92) + 'T12:00:00Z', summit_at: ago(80) });
  w.masterLog.rows = [ML_HEADERS.slice(),
    ml('E1', '8012345', 'Quill, Pat', 'Disruption', ago(150)), ml('E2', '8012345', 'Quill, Pat', 'Disruption', ago(120)), ml('E3', '8012345', 'Quill, Pat', 'Disruption', ago(101), 'Admin Lunch Detention (LD)'),
    ml('E4', '8012345', 'Quill, Pat', 'Technology', ago(40))];
  w.user.email = 'teacher@easdpa.org';
  assert.ok(/panel leads/.test(w.call({ action: 'trailbackOutcomes', payload: {} }).error));
  w.user.email = 'lead@easdpa.org';
  const r = w.call({ action: 'trailbackOutcomes', payload: {} });
  assert.ok(r.ok, r.error);
  const d = r.data, o = d.students[0];
  assert.strictEqual(d.overall.students, 1);
  assert.strictEqual(d.pipeline.completed, 1);
  assert.strictEqual(d.pipeline.detentionsCleared, 1);
  assert.strictEqual(d.pipeline.summits, 1);
  assert.strictEqual(o.after, 1);
  assert.strictEqual(o.sameAfter, 0, 'no Disruption after the panel');
  assert.strictEqual(o.daysToFinish, 8);
  assert.ok(o.perMonthAfter < o.perMonthBefore);
  assert.deepStrictEqual(J(o.windows[30]), { before: 2, after: 0, sameBefore: 2, sameAfter: 0 }, 'E2 and E3 fall in the 30 days before');
  assert.strictEqual(d.overall.windows[90].noRepeat, 100);
  assert.strictEqual(d.byBehavior[0].label, 'Disruption');
});

console.log('\n' + passed + ' tests passed');
