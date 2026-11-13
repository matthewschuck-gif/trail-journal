// Offline checks for the Trailback Panel sign-up (fake Google services, real .gs).
//   node tests/panel-signup.test.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const dir = path.join(__dirname, '..', 'apps-script');
let passed = 0;
function test(name, fn) { fn(); passed++; console.log('ok - ' + name); }

const OWNER = 'matthew_schuck@easdpa.org';
function world(opts) {
  opts = opts || {};
  const mails = [], triggers = [];
  const settings = Object.assign({ SITE_URL: 'https://trailjournal.org', PANEL_ADMINS: '', PANEL_QUESTION_SHEET_URL: 'https://example.org/sheet' }, opts.settings);
  const fmt = (d, tz, f) => {
    const p = new Intl.DateTimeFormat('en-CA', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hour12: false }).formatToParts(d)
      .reduce((o, x) => { o[x.type] = x.value; return o; }, {});
    if (f === 'yyyy-MM-dd') return p.year + '-' + p.month + '-' + p.day;
    if (f === 'HH:mm') return (p.hour === '24' ? '00' : p.hour) + ':' + p.minute;
    return d.toISOString();
  };
  let uuid = 0;
  const ctx = { console, JSON, Math, Date, Object, Array, String, Number, Intl,
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: () => null }) },
    Session: { getActiveUser: () => ({ getEmail: () => '' }), getEffectiveUser: () => ({ getEmail: () => OWNER }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    MailApp: { sendEmail: o => mails.push(o) },
    Utilities: { formatDate: fmt, getUuid: () => 'id' + (++uuid), newBlob: (text, type, name) => ({ text, type, name }) },
    ScriptApp: { getProjectTriggers: () => triggers, deleteTrigger() {},
      newTrigger: fn => { const b = { timeBased: () => b, everyDays: () => b, atHour: h => { b.hour = h; return b; }, inTimezone: () => b, create: () => { triggers.push({ fn, hour: b.hour, getHandlerFunction: () => fn }); } }; return b; } },
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readdirSync(dir).filter(f => f.endsWith('.gs')).sort().map(f => fs.readFileSync(path.join(dir, f), 'utf8')).join('\n'), ctx);
  const tabs = { panel_cases: [], panel_signups: [] };
  ctx.panelTable_ = tab => ({ rows: tabs[tab], add: o => { tabs[tab].push(o); return o; }, set: (row, patch) => Object.assign(row, patch) });
  ctx.webSetting_ = k => settings[k] || '';
  ctx.panelToday_ = () => opts.today || '2026-11-13';
  ctx.safeAlert_ = () => {};
  const journals = opts.journals || [];
  ctx.recordsTable_ = () => ({ values: journals, row: i => journals[i] });
  ctx.rowCreatedMs_ = r => Date.parse(r.created_at);
  const as = email => body => ctx.dispatch_(Object.assign({ _staffEmail: email }, body));
  return { ctx, tabs, mails, triggers, as, owner: as(OWNER) };
}
const CASE = { studentName: 'Ann Lee', studentId: '8001111', grade: '7', behavior: 'Repeated tardiness', panelDate: '2026-11-18', startTime: '9:50', minutes: 20, seats: 2, room: 'Room 12' };

test('only signed-in staff reach the sign-up, and only panel leads add cases', () => {
  const w = world();
  assert.throws(() => w.ctx.dispatch_({ action: 'panelBoard' }), /sign in/i);
  assert.throws(() => w.ctx.dispatch_({ action: 'panelBoard', adminPassword: 'anything' }), /sign in/i, 'the shared password is not an identity');
  assert.throws(() => w.as('teacher@easdpa.org')({ action: 'panelSaveCase', payload: CASE }), /panel leads/);
  w.owner({ action: 'panelSaveCase', payload: CASE });
  assert.strictEqual(w.tabs.panel_cases.length, 1);
  assert.strictEqual(w.tabs.panel_cases[0].start_time, '09:50');
  const leads = world({ settings: { PANEL_ADMINS: 'kelly_rigg@easdpa.org' } });
  leads.as('kelly_rigg@easdpa.org')({ action: 'panelSaveCase', payload: CASE });
  assert.throws(() => leads.owner({ action: 'panelSaveCase', payload: CASE }), /panel leads/, 'a set list replaces the owner default');
});

test('signing up takes a seat, sends one calendar invite, and a full or repeated sign-up is refused', () => {
  const w = world();
  const id = w.owner({ action: 'panelSaveCase', payload: CASE }).id;
  const r = w.as('austin_brass@easdpa.org')({ action: 'panelSignUp', payload: { caseId: id } });
  assert.strictEqual(r.invite, 'sent');
  assert.strictEqual(w.mails.length, 1);
  const m = w.mails[0];
  assert.strictEqual(m.to, 'austin_brass@easdpa.org');
  assert.ok(/Ann Lee/.test(m.subject));
  const ics = m.attachments[0].text;
  assert.ok(/DTSTART;TZID=America\/New_York:20261118T095000/.test(ics) && /DTEND;TZID=America\/New_York:20261118T101000/.test(ics), ics);
  assert.ok(/UID:trailback-id1-austinbrasseasdpaorg@trailjournal\.org/.test(ics));
  assert.throws(() => w.as('austin_brass@easdpa.org')({ action: 'panelSignUp', payload: { caseId: id } }), /already on this panel/);
  w.as('maggie_myers@easdpa.org')({ action: 'panelSignUp', payload: { caseId: id, name: 'Maggie Myers' } });
  assert.ok(/Austin Brass/.test(w.mails[1].htmlBody), 'the second invite names who is already on the panel');
  assert.throws(() => w.as('third@easdpa.org')({ action: 'panelSignUp', payload: { caseId: id } }), /full/);
  const board = w.as('austin_brass@easdpa.org')({ action: 'panelBoard' });
  assert.deepStrictEqual([board.cases[0].seatsLeft, board.cases[0].mine, board.cases[0].panelists], [0, true, ['Austin Brass', 'Maggie Myers']]);
  assert.strictEqual(board.isAdmin, false);
});

test('stepping off frees the seat and sends a cancellation for the same calendar entry', () => {
  const w = world();
  const id = w.owner({ action: 'panelSaveCase', payload: CASE }).id;
  w.as('a@easdpa.org')({ action: 'panelSignUp', payload: { caseId: id } });
  const r = w.as('a@easdpa.org')({ action: 'panelWithdraw', payload: { caseId: id } });
  assert.strictEqual(r.cancel, 'sent');
  const cancel = w.mails[1].attachments[0];
  assert.ok(/method=CANCEL/.test(cancel.type) && /STATUS:CANCELLED/.test(cancel.text));
  assert.ok(cancel.text.includes(w.mails[0].attachments[0].text.match(/UID:[^\r]+/)[0]), 'same UID as the invite');
  assert.strictEqual(w.as('a@easdpa.org')({ action: 'panelBoard' }).cases[0].seatsLeft, 2);
  assert.throws(() => w.as('a@easdpa.org')({ action: 'panelWithdraw', payload: { caseId: id } }), /not on this panel/);
});

test('the case brief carries the Trail Journal and the Climb-er summary, goes out at 3:00 the weekday before, once', () => {
  const brief = { at: '2026-11-13T14:00:00Z', profile: { name: 'Ann Lee', team: 'M', squad: 'Gold', homeroom: 'Smith', ceGroup: 'Math B' },
    pattern: { days: 42, total: 45, counts: [{ label: 'Incidents', recent: 3, year: 5 }], months: [{ month: '2026-09', count: 2 }, { month: '2026-11', count: 3 }],
      records: [{ date: '11/10/2026', source: 'Incident reports', headline: 'Late to class <b>again</b>' }, { date: '9/8/2026', source: 'Incident reports', headline: 'First week of school' }] },
    strengths: { counts: [{ label: 'Shoutouts', year: 2 }], recent: [{ date: '10/2/2026', source: 'Cairn shoutouts received', headline: 'Helped a classmate' }], activities: ['Soccer'] } };
  const w = world({ today: '2026-11-17', journals: [
    { student_id: '8001111', student_name: 'Ann Lee', status: 'in_progress', created_at: '2026-11-01T12:00:00Z', reflection_text: '=== OLD ===\nolder one' },
    { student_id: '8001111', student_name: 'Ann Lee', status: 'submitted', created_at: '2026-11-12T12:00:00Z', reflection_text: '=== WHAT HAPPENED ===\nI was late because I stopped at my locker.' },
    { student_id: '8002222', student_name: 'Bo Kim', status: 'submitted', created_at: '2026-11-13T12:00:00Z', reflection_text: 'someone else' }] });
  const id = w.owner({ action: 'panelSaveCase', payload: CASE }).id;
  w.owner({ action: 'panelSaveCase', payload: Object.assign({}, CASE, { studentName: 'Later Case', studentId: '8003333', panelDate: '2026-11-25' }) });
  w.tabs.panel_cases[0].climber_brief = JSON.stringify(brief);
  w.as('a@easdpa.org')({ action: 'panelSignUp', payload: { caseId: id } });
  w.as('b@easdpa.org')({ action: 'panelSignUp', payload: { caseId: id } });
  w.mails.length = 0;
  assert.strictEqual(w.ctx.panelDailyBriefs(), 2);
  assert.deepStrictEqual(w.mails.map(m => m.to), ['a@easdpa.org', 'b@easdpa.org']);
  const html = w.mails[0].htmlBody;
  assert.ok(/stopped at my locker/.test(html) && !/older one/.test(html) && !/someone else/.test(html), 'the newest journal for this student only');
  assert.ok(/Gold squad/.test(html) && /Team M/.test(html) && /Core Extension: Math B/.test(html));
  assert.ok(/Pattern this school year/.test(html) && /Incidents: 5 this year \(3 in the last 42 days\)/.test(html) && /Shoutouts: 2 this year/.test(html) && /Soccer/.test(html));
  assert.ok(/Sep 2 · Nov 3/.test(html) && /First week of school/.test(html) && /newest 2 of 45 records/.test(html), 'the whole year: months, September record, and a note when the list is cut');
  assert.ok(html.includes('Late to class &lt;b&gt;again&lt;/b&gt;'), 'record text is escaped');
  assert.strictEqual(w.ctx.panelDailyBriefs(), 0, 'not sent twice');
  // Someone who signs up after the brief went out gets it with their invite.
  w.tabs.panel_cases[0].seats = 3;
  w.mails.length = 0;
  w.as('late@easdpa.org')({ action: 'panelSignUp', payload: { caseId: id } });
  assert.deepStrictEqual(w.mails.map(m => /Case brief/.test(m.subject)), [false, true]);
  // Friday looks ahead to Monday.
  assert.strictEqual(w.ctx.panelNextWeekday_('2026-11-13'), '2026-11-16');
});

test('a case with no Climb-er summary or Trail Journal says so; leads can send the brief by hand; others cannot', () => {
  const w = world();
  const id = w.owner({ action: 'panelSaveCase', payload: Object.assign({}, CASE, { studentId: '' }) }).id;
  w.as('a@easdpa.org')({ action: 'panelSignUp', payload: { caseId: id } });
  assert.throws(() => w.as('a@easdpa.org')({ action: 'panelSendBrief', payload: { caseId: id } }), /panel leads/);
  w.mails.length = 0;
  assert.strictEqual(w.owner({ action: 'panelSendBrief', payload: { caseId: id } }).sent, 1);
  assert.ok(/not on this case yet/.test(w.mails[0].htmlBody) && /No Trail Journal found/.test(w.mails[0].htmlBody));
  const board = w.owner({ action: 'panelBoard' });
  assert.ok(board.isAdmin && board.cases[0].briefSentAt && board.cases[0].journal === null && board.cases[0].climberReady === false);
});

test('closed cases are hidden from staff, past cases from everyone, and the 3:00 trigger is set once', () => {
  const w = world();
  const id = w.owner({ action: 'panelSaveCase', payload: CASE }).id;
  w.owner({ action: 'panelSaveCase', payload: Object.assign({}, CASE, { panelDate: '2026-11-01' }) });
  w.owner({ action: 'panelSaveCase', payload: Object.assign({}, CASE, { id, status: 'closed' }) });
  assert.strictEqual(w.as('a@easdpa.org')({ action: 'panelBoard' }).cases.length, 0);
  assert.strictEqual(w.owner({ action: 'panelBoard' }).cases.length, 1);
  assert.throws(() => w.as('a@easdpa.org')({ action: 'panelSignUp', payload: { caseId: id } }), /closed/);
  w.ctx.setupPanelBriefTrigger();
  assert.deepStrictEqual(w.triggers.map(t => [t.fn, t.hour]), [['panelDailyBriefs', 15]]);
});

console.log(`\n${passed} tests passed`);
