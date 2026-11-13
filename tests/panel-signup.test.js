// Offline checks for the Trailback Panel sign-up (13_PanelSignup.gs) on the live case list.
// Fake Google services, real .gs code, made-up students.
//   node tests/panel-signup.test.js
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

const aCase = { student_id: '8012345', student_name: 'Pat Quill', grade: '8', squad: 'Gold', behavior_type: 'Disrespect',
  consequence_type: 'After School Detention', consequence_original_date: '2026-12-01', master_log_row: 'EMS-20261201-0042',
  referring_staff: 'Ms. AP', incident_summary: 'Argued with a sub and left class.', status: 'offered', source: 'admin_conference' };

// A case Scheduled for a Wednesday a week or more out.
function scheduled(w, extra) {
  const c = w.call({ action: 'createTrailbackCase', payload: Object.assign({}, aCase, extra || {}) }).data;
  w.step(c.case_id, 'consent');
  const wed = wednesdayAfter(7);
  const r = w.step(c.case_id, 'schedule', { date: wed });
  assert.ok(r.ok, r.error);
  return { id: c.case_id, wed };
}
const as = (w, email) => body => { w.user.email = email; return w.call(body); };

test('the sign-up page shows Scheduled cases from the live case list, with the Settings defaults', () => {
  const w = world();
  const s = scheduled(w);
  w.call({ action: 'createTrailbackCase', payload: Object.assign({}, aCase, { student_id: '8099999', master_log_row: 'EMS-20261201-0043' }) }); // offered only
  const b = w.call({ action: 'panelBoard', payload: {} });
  assert.ok(b.ok, b.error);
  assert.strictEqual(b.data.cases.length, 1, 'only the scheduled case');
  const c = b.data.cases[0];
  assert.strictEqual(c.id, s.id);
  assert.strictEqual(c.panelDate, s.wed);
  assert.strictEqual(c.studentName, 'Pat Quill');
  assert.strictEqual(c.startTime, '12:00');
  assert.strictEqual(c.minutes, 25);
  assert.strictEqual(c.seats, 6);
  assert.strictEqual(c.status, 'open');
  w.user.email = '';
  assert.ok(/sign in/i.test(w.call({ action: 'panelBoard', payload: {} }).error), 'staff sign-in only');
});

test('only panel leads set panel details, and only the details (never the student)', () => {
  const w = world({ settings: [['PANEL_ADMINS', 'lead@easdpa.org', '']] });
  const s = scheduled(w);
  const teacher = as(w, 'teacher@easdpa.org'), lead = as(w, 'lead@easdpa.org');
  assert.ok(/panel leads/.test(teacher({ action: 'panelSaveCase', payload: { id: s.id, startTime: '9:50' } }).error));
  const r = lead({ action: 'panelSaveCase', payload: { id: s.id, startTime: '9:50', minutes: 30, seats: 2, room: 'Room 12', notes: 'Bring water', studentName: 'Someone Else' } });
  assert.ok(r.ok, r.error);
  const c = lead({ action: 'panelBoard', payload: {} }).data.cases[0];
  assert.deepStrictEqual([c.startTime, c.minutes, c.seats, c.room, c.notes, c.studentName], ['09:50', 30, 2, 'Room 12', 'Bring water', 'Pat Quill']);
  const hist = w.call({ action: 'trailbackCase', payload: { case_id: s.id } }).data.history;
  assert.ok(/Panel details: 09:50/.test(hist[hist.length - 1].note), 'the change is on the case history');
  lead({ action: 'panelSaveCase', payload: { id: s.id, startTime: '9:50', status: 'closed' } });
  assert.strictEqual(teacher({ action: 'panelBoard', payload: {} }).data.cases.length, 0, 'closed cases are hidden from non-leads');
});

test('signing up sends a calendar invite at once; seats fill; stepping off sends a cancellation', () => {
  const w = world({ settings: [['PANEL_ADMINS', 'lead@easdpa.org', '']] });
  const s = scheduled(w);
  as(w, 'lead@easdpa.org')({ action: 'panelSaveCase', payload: { id: s.id, startTime: '12:30', seats: 2, room: 'Library' } });
  const t1 = as(w, 'ann_one@easdpa.org'), t2 = as(w, 'bo_two@easdpa.org'), t3 = as(w, 'cy_three@easdpa.org');
  w.sent.length = 0;
  const r = t1({ action: 'panelSignUp', payload: { caseId: s.id } });
  assert.ok(r.ok, r.error);
  assert.strictEqual(r.data.invite, 'sent');
  const inv = w.sent[0];
  assert.strictEqual(inv.to, 'ann_one@easdpa.org');
  assert.ok(/Trailback Panel: Pat Quill/.test(inv.subject));
  assert.ok(inv.attachments && /BEGIN:VCALENDAR/.test(inv.attachments[0].text) && /METHOD:REQUEST|VEVENT/.test(inv.attachments[0].text), 'a calendar invite is attached');
  assert.ok(/Library/.test(inv.htmlBody));
  assert.ok(/already on this panel/.test(t1({ action: 'panelSignUp', payload: { caseId: s.id } }).error));
  t2({ action: 'panelSignUp', payload: { caseId: s.id } });
  assert.ok(/full/.test(t3({ action: 'panelSignUp', payload: { caseId: s.id } }).error), 'two seats');
  const board = t3({ action: 'panelBoard', payload: {} }).data.cases[0];
  assert.deepStrictEqual(J(board.panelists), ['Ann One', 'Bo Two']);
  assert.strictEqual(board.seatsLeft, 0);
  w.sent.length = 0;
  const off = t1({ action: 'panelWithdraw', payload: { caseId: s.id } });
  assert.ok(off.ok && off.data.cancel === 'sent');
  assert.ok(/STATUS:CANCELLED/.test(w.sent[0].attachments[0].text));
  assert.ok(t3({ action: 'panelSignUp', payload: { caseId: s.id } }).ok, 'the open seat can be taken');
});

test('the brief goes out the school day before: signed-up panelists get the case brief with Climb-er + Trail Journal; empty cases go to the Panel Members; never twice', () => {
  const w = world({ settings: [['TRAILBACK_EMAILS', 'ON', ''], ['PANEL_ADMINS', 'lead@easdpa.org', '']] });
  const s = scheduled(w);
  const other = scheduled(w, { student_id: '8099999', student_name: 'Lee Rowe', master_log_row: 'EMS-20261201-0043' });
  w.tabs['Panel Members'] = makeSheet([['Name', 'Email', 'Role', 'Active'], ['Kelly', 'kelly@easdpa.org', 'Lead', 'TRUE'], ['Sam', 'sam@easdpa.org', 'Student', 'TRUE']]);
  as(w, 'ann_one@easdpa.org')({ action: 'panelSignUp', payload: { caseId: s.id } });
  // Climb-er wrote its summary onto the case row.
  w.genericUpdate_('panel_cases', { case_id: s.id }, { climber_brief: JSON.stringify({ at: '2026-11-17T10:00:00Z', profile: { squad: 'Gold', team: 'Summit' },
    pattern: { days: 42, total: 1, counts: [{ label: 'Office referrals', year: 3, recent: 1 }], months: [{ month: '2026-10', count: 3 }], records: [{ date: '10/2', source: 'Master Log', headline: 'Disrespect' }] },
    strengths: { counts: [{ label: 'Shout-outs', year: 4 }], recent: [], activities: ['Robotics'] } }) });
  const brief = w.panelSchoolDayBefore_(s.wed);
  w.todayYmd_ = () => w.ymdShift_ ? w.ymdShift_(brief, -1) : '2000-01-01';
  w.sent.length = 0;
  w.todayYmd_ = () => '2000-01-01';
  assert.strictEqual(J(w.panelDailyBriefs_(false)).briefs.length, 0, 'too early: nothing');
  w.todayYmd_ = () => brief;
  const r = J(w.panelDailyBriefs_(false));
  assert.strictEqual(r.briefs.length, 1);
  const mail = w.sent.find(m => m.to === 'ann_one@easdpa.org');
  assert.ok(mail && /Case brief/.test(mail.subject));
  assert.ok(/Robotics/.test(mail.htmlBody) && /Office referrals: 3 this year/.test(mail.htmlBody) && /Summit/.test(mail.htmlBody), 'the Climb-er summary is in it');
  const fb = w.sent.find(m => /sign-ups needed/.test(m.subject));
  assert.ok(fb && fb.to === 'kelly@easdpa.org', 'the empty case goes to the adult Panel Members');
  assert.ok(/Lee Rowe/.test(fb.htmlBody) && /sign-up page/.test(fb.htmlBody));
  const n = w.sent.length;
  w.panelDailyBriefs_(false);
  w.trailbackTuesdayBrief(); // the old Tuesday timer runs the same thing
  assert.strictEqual(w.sent.length, n, 'never twice');
  // Signing up after the brief went out: the brief comes with the invite.
  w.sent.length = 0;
  as(w, 'bo_two@easdpa.org')({ action: 'panelSignUp', payload: { caseId: s.id } });
  assert.ok(w.sent.some(m => m.to === 'bo_two@easdpa.org' && /Case brief/.test(m.subject)));
  // TRAILBACK_EMAILS OFF = nothing automatic.
  const off = world();
  assert.ok(/OFF/.test(off.panelDailyBriefs_(false).skipped));
  void other;
});

test('the timers: one daily 3:00 brief replaces the Tuesday one; the days off are skipped when counting back', () => {
  const w = world();
  w.setupTrailbackTriggers();
  assert.ok(w.triggers.includes('panelDailyBriefs') && !w.triggers.includes('trailbackTuesdayBrief'));
  w.tabs['No School Dates'] = makeSheet([['Date'], ['2026-11-24']]);
  assert.strictEqual(w.panelSchoolDayBefore_('2026-11-25'), '2026-11-23', 'Tuesday off: the brief goes Monday');
  assert.strictEqual(w.panelSchoolDayBefore_('2026-11-30'), '2026-11-27', 'a Monday panel: the Friday before');
});

console.log('\n' + passed + ' tests passed');
