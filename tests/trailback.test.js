// Offline checks for the Trailback Panel rules (11_Trailback.gs, 07_ClashOfClasses.gs, 02_AIProxy.gs).
// Fake Google services, real .gs code, made-up students.
//   node tests/trailback.test.js
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
    getDataRange() { return { getValues: () => s.rows.map(r => r.slice()) }; },
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
    setFrozenRows() {},
  };
  return s;
}

function world(opts) {
  opts = opts || {};
  const tabs = {}, clashEvents = makeSheet([['id', 'name', 'type', 'date', 'note', 'multiplier', 'BLACK', 'GOLD', 'GREY', 'PURPLE']]);
  const user = { email: opts.email === undefined ? 'ap@easdpa.org' : opts.email };
  const ss = { getSheetByName: n => tabs[n] || null, insertSheet: n => (tabs[n] = makeSheet([])) };
  const prompts = [];
  const ctx = { console, JSON, Math, Date, Object, Array, String, Number,
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, openById: () => ({ getSheetByName: n => n === 'events' ? clashEvents : null }), getUi: () => { throw new Error('no ui'); } },
    Session: { getActiveUser: () => ({ getEmail: () => user.email }), getEffectiveUser: () => ({ getEmail: () => 'me@easdpa.org' }) },
    Utilities: { formatDate: () => '2026-12-09', getUuid: () => 'u' + Math.random().toString(36).slice(2) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    UrlFetchApp: { fetch: (u, o) => { prompts.push(JSON.parse(o.payload)); return { getResponseCode: () => 200, getContentText: () => JSON.stringify({ content: [{ text: '{"analysis":"ok","recommendation":"ok"}' }] }) }; } },
    Logger: { log() {} },
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readdirSync(dir).filter(f => f.endsWith('.gs')).sort().map(f => fs.readFileSync(path.join(dir, f), 'utf8')).join('\n') + ';this.SCHEMA = SCHEMA; this.TABS = TABS;', ctx);
  // The live Sheet today: panel_sessions has every column up to student_id, nothing newer.
  const ps = ctx.SCHEMA.panel_sessions;
  tabs.panel_sessions = makeSheet([ps.slice(0, ps.indexOf('student_id') + 1)]);
  tabs.Settings = makeSheet([['Key', 'Value', 'Notes']].concat(opts.settings || []));
  if (opts.noSchool) tabs['No School Dates'] = makeSheet([['Date', 'Note']].concat(opts.noSchool.map(d => [d, ''])));
  ctx.getProp_ = k => ({ APP_TOKEN: 'tok', ADMIN_PASSWORD: 'staffpw', ANTHROPIC_API_KEY: 'k' })[k];
  // A page call on the staff link: tjCall puts the signed-in account on the body.
  ctx.call = body => J(JSON.parse(ctx.tjCall(JSON.stringify(body))));
  ctx.tabs = tabs; ctx.clashEvents = clashEvents; ctx.user = user; ctx.prompts = prompts;
  return ctx;
}

const record = extra => Object.assign({
  session_id: 'p1', panel_date: '2026-12-09', student_name: 'Pat Quill', student_initials: 'P.Q.', student_id: '8012345',
  grade: '8', recommendation_narrative: 'Repair the hallway display with the art teacher.', clash_squad: 'Gold',
  followup_dates: { checkin_3: '2026-12-14', checkin_10: '2026-12-23', checkin_30: '2027-01-29' },
}, extra || {});

// ── initials ──
test('initials: first and last initial only', () => {
  const w = world();
  assert.strictEqual(w.initialsOf_('Pat Quill'), 'P.Q.');
  assert.strictEqual(w.initialsOf_('  mary-kate  van der berg '), 'M.B.');
  assert.strictEqual(w.initialsOf_('Cher'), 'C.');
  assert.strictEqual(w.initialsOf_(''), '');
});

// ── school days ──
test('school days: a Wednesday panel counts from Thursday and skips the weekend', () => {
  const w = world();
  assert.strictEqual(w.schoolDaysAfter_('2026-12-09', 3, {}), '2026-12-14'); // Thu, Fri, Mon
  assert.strictEqual(w.schoolDaysAfter_('2026-12-09', 10, {}), '2026-12-23');
});

test('school days: the No School Dates tab is skipped (winter break), in any date format', () => {
  const w = world({ noSchool: ['2026-12-24', '12/25/2026', new Date(2026, 11, 28), '2026-12-29', '2026-12-30', '2026-12-31', '2027-01-01'] });
  const d = w.followupSchedule_({ from: '2026-12-09' });
  assert.strictEqual(d.checkin_3, '2026-12-14');
  assert.strictEqual(d.checkin_10, '2026-12-23');
  // 30 school days: Dec 10-23 (10), break skipped, then Jan 4 onward (20 more) = Jan 29
  assert.strictEqual(d.checkin_30, '2027-01-29');
  const noBreak = world().followupSchedule_({ from: '2026-12-09' });
  assert.strictEqual(noBreak.checkin_30, '2027-01-20');
});

test('followupDates is staff-only and returns the three check-in dates', () => {
  const w = world({ email: '' });
  assert.ok(/staff password/i.test(w.call({ action: 'followupDates', payload: { from: '2026-12-09' } }).error));
  const ok = world().call({ action: 'followupDates', payload: { from: '2026-12-09' } });
  assert.deepStrictEqual([ok.data.checkin_3, ok.data.checkin_10], ['2026-12-14', '2026-12-23']);
});

test('old rows keep reading: two_days/one_week/... still come back with their old labels', () => {
  const w = world();
  assert.deepStrictEqual(J(w.panelCheckins_({ two_days: '2026-10-03', one_month: '2026-11-01' })).map(c => c.label + '=' + c.date),
    ['2 days=2026-10-03', '1 week=', '2 weeks=', '1 month=2026-11-01']);
  assert.deepStrictEqual(J(w.panelCheckins_({ checkin_3: 'a', checkin_10: 'b', checkin_30: 'c' })).map(c => c.label),
    ['3 school days', '10 school days', '30 school days']);
});

// ── AI: no names ──
test('panel AI helpers get initials, never the name (even from an old page that sends the name)', () => {
  const w = world();
  w.call({ action: 'ai', type: 'panel_inquiry_analysis', payload: { initials: 'P.Q.', incident: 'x', inquiryNotes: [] } });
  w.call({ action: 'ai', type: 'panel_recommendation_draft', payload: { initial: 'Pat Quill', incident: 'x' } });
  assert.strictEqual(w.prompts.length, 2);
  const text = JSON.stringify(w.prompts);
  assert.ok(!/Pat Quill|Quill/.test(text), 'name reached the AI: ' + text.slice(0, 300));
  assert.ok(/Student initials: P\.Q\./.test(w.prompts[0].messages[0].content));
  assert.ok(/Student initials: P\.Q\./.test(w.prompts[1].messages[0].content));
});

// ── approval ──
test('approval needs a Google sign-in on the staff link; the staff password alone is not enough', () => {
  const w = world({ email: '' });
  const r = w.call({ action: 'approvePanelPlan', adminPassword: 'staffpw', payload: { record: record() } });
  assert.ok(/Google sign-in/.test(r.error), JSON.stringify(r));
});

test('approval records the signed-in account and time (not what the page sends), status heard_approved', () => {
  const w = world();
  const r = w.call({ action: 'approvePanelPlan', payload: { record: record({ approved_by: 'someone-else@x', approved_at: '1999' }) } });
  assert.strictEqual(r.data.approved_by, 'ap@easdpa.org');
  assert.ok(r.data.approved_at && r.data.approved_at !== '1999');
  const row = w.genericQuery_('panel_sessions', { session_id: 'p1' })[0];
  assert.strictEqual(row.status, 'heard_approved');
  assert.strictEqual(row.approved_by, 'ap@easdpa.org');
  assert.strictEqual(row.student_name, 'Pat Quill');
  assert.ok(w.tabs.panel_sessions.rows[0].indexOf('approved_by') > -1, 'new columns added to the live tab');
});

test('PANEL_APPROVERS limits who can approve', () => {
  const w = world({ email: 'teacher@easdpa.org', settings: [['PANEL_APPROVERS', 'ap@easdpa.org, kelly@easdpa.org', '']] });
  assert.ok(/not on the PANEL_APPROVERS/.test(w.call({ action: 'approvePanelPlan', payload: { record: record() } }).error));
  w.user.email = 'kelly@easdpa.org';
  assert.strictEqual(w.call({ action: 'approvePanelPlan', payload: { record: record() } }).data.approved_by, 'kelly@easdpa.org');
});

test('submit refuses an unapproved session; after approval it saves as repair_in_progress and keeps the approval', () => {
  const w = world();
  assert.ok(/not been approved/.test(w.call({ action: 'submitPanelSession', payload: { record: record() } }).error));
  w.call({ action: 'approvePanelPlan', payload: { record: record() } });
  const r = w.call({ action: 'submitPanelSession', payload: { record: record({ followup_person: 'Ms. Ray', approved_by: 'forged@x', status: 'closed_completed',
    fu_criteria: 'Apology done and three days of lunch help.', fu_fallback: 'The detention is served.' }) } });
  assert.strictEqual(r.data.status, 'repair_in_progress');
  assert.strictEqual(r.data.fu_criteria, 'Apology done and three days of lunch help.', 'completion criteria saved (it used to be dropped)');
  assert.strictEqual(r.data.fu_fallback, 'The detention is served.');
  assert.strictEqual(r.data.approved_by, 'ap@easdpa.org');
  assert.strictEqual(r.data.followup_person, 'Ms. Ray');
  assert.strictEqual(w.tabs.panel_sessions.rows.length, 2, 'one row per session');
});

test('the generic insert/update cannot write an approval', () => {
  const w = world();
  w.call({ action: 'insert', table: 'panel_sessions', row: record({ session_id: 'p2', approved_by: 'forged@x', status: 'heard_approved' }) });
  const row = w.genericQuery_('panel_sessions', { session_id: 'p2' })[0];
  assert.strictEqual(row.approved_by, '');
  assert.notStrictEqual(row.status, 'heard_approved');
  w.call({ action: 'update', table: 'panel_sessions', match: { session_id: 'p2' }, patch: { approved_by: 'forged@x' } });
  assert.strictEqual(w.genericQuery_('panel_sessions', { session_id: 'p2' })[0].approved_by, '');
  assert.ok(/not been approved/.test(w.call({ action: 'submitPanelSession', payload: { record: record({ session_id: 'p2' }) } }).error));
});

// ── Clash ──
test('Clash: a completed plan makes one pending request with no name or initials; asking twice for a case gives the same request', () => {
  const w = world();
  const a = J(w.requestClashPoint_({ squad: 'Gold', caseId: 'c1', requestedBy: 'kelly@easdpa.org' }));
  const b = J(w.requestClashPoint_({ squad: 'Gold', caseId: 'c1' }));
  assert.strictEqual(a.id, b.id);
  assert.strictEqual(a.status, 'pending');
  assert.strictEqual(a.note, 'Trailback plan completed.');
  assert.strictEqual(w.clashEvents.rows.length, 1, 'nothing goes to the public sheet before approval');
  const queue = w.call({ action: 'clashRequests' }).data;
  assert.strictEqual(queue.length, 1);
  assert.ok(!/Pat|Quill|P\.Q\./.test(JSON.stringify(queue)));
});

test('Clash: approval appends one anonymous trailback_panel row; the same request cannot award twice', () => {
  const w = world({ settings: [['TRAILBACK_CLASH_POINTS', '2', '']] });
  const req = J(w.requestClashPoint_({ squad: 'Purple', caseId: 'c9' }));
  const r = w.call({ action: 'decideClashRequest', payload: { id: req.id, decision: 'approve' } });
  assert.strictEqual(r.data.status, 'approved');
  assert.deepStrictEqual(J(w.clashEvents.rows[1]), ['', 'Trailback Panel', 'trailback_panel', '2026-12-09', 'Trailback plan completed.', '', '', '', '', 2]);
  assert.ok(/already approved/.test(w.call({ action: 'decideClashRequest', payload: { id: req.id, decision: 'approve' } }).error));
  assert.strictEqual(w.clashEvents.rows.length, 2);
});

test('Clash: the queue and decisions are staff-only; a declined request adds nothing', () => {
  const w = world({ email: '' });
  assert.ok(/staff password/i.test(w.call({ action: 'clashRequests' }).error));
  const req = J(w.requestClashPoint_({ squad: 'Grey', caseId: 'c3' }));
  assert.ok(/staff password/i.test(w.call({ action: 'decideClashRequest', payload: { id: req.id, decision: 'approve' } }).error));
  w.user.email = 'ap@easdpa.org';
  assert.strictEqual(w.call({ action: 'decideClashRequest', payload: { id: req.id, decision: 'decline' } }).data.status, 'declined');
  assert.strictEqual(w.clashEvents.rows.length, 1);
});

test('Panel page no longer talks to the old Clash Supabase project or sends the name to the AI', () => {
  const src = fs.readFileSync(path.join(dir, 'Panel.html'), 'utf8');
  assert.ok(!/yvutgnpgnthvnonhgdux|COC_ANON|awardPanelCoCPoint|supabase\.co/i.test(src));
  assert.ok(!/P\.initial\b|s-initial/.test(src));
  assert.ok(!/status:\s*'complete'/.test(src));
});

test('Panel preview (?preview=true) answers every write itself: approval, submit, back-write, emails', () => {
  const src = fs.readFileSync(path.join(dir, 'Panel.html'), 'utf8');
  const m = src.match(/const PREVIEW_WRITES = (\[[^\]]*\]);/);
  assert.ok(m, 'PREVIEW_WRITES list');
  const writes = JSON.parse(m[1].replace(/'/g, '"'));
  ['approvePanelPlan', 'submitPanelSession', 'insert', 'update', 'sendFollowupEmails'].forEach(a => assert.ok(writes.includes(a), a));
  // The check sits at the top of callBackend, before anything is sent.
  const cb = src.slice(src.indexOf('async function callBackend('));
  assert.ok(cb.indexOf('PREVIEW_WRITES') < cb.indexOf('tjCall_'));
  // The gate's auto-unlock comes after ADMIN_PW_VALUE is declared (it threw before).
  assert.ok(src.indexOf('let ADMIN_PW_VALUE') < src.indexOf('if (_savedPanelPw) ADMIN_PW_VALUE'));
});

console.log('\n' + passed + ' tests passed');
