/**
 * Trail Journal -- Apps Script backend
 * File 18: Trailback outcomes (?page=outcomes on the staff link; panel leads and the coordinator).
 *
 * How do students who went through a Trailback Panel do the rest of the year?
 *   - The cases: how many reached a panel, finished the plan (and the detention was cleared),
 *     reverted or declined; days from the panel to a finished plan; check-ins on time; Summits.
 *   - Each student's Master Log record (the Accountability Sheet) before and after their panel:
 *     entries per month in the time before the panel vs. the time since; the same behavior again
 *     in the 30, 60 and 90 days after vs. the same days before; detentions after.
 *   - All of that by behavior type and by the panel's need (Connection, Awareness, Skills,
 *     Regulation), and a row per student for staff.
 * Counts come from the Master Log only (consequences); the months are school-year months.
 */

const OUTCOME_WINDOWS_ = [30, 60, 90];

function canSeeOutcomes_(email) {
  const e = String(email || '').toLowerCase();
  return !!e && (panelIsAdmin_(e) || coordinatorEmails_().toLowerCase().split(',').indexOf(e) > -1);
}

// { 800#: [{ day, type, cons }] } from the Master Log (cancelled / rescheduled rows left out).
function masterLogByStudent_() {
  const id = webSetting_('ACCOUNTABILITY_SHEET_ID');
  if (!id) return {};
  const sh = SpreadsheetApp.openById(id).getSheetByName(ML_TAB_);
  if (!sh || sh.getLastRow() < 2) return {};
  const v = sh.getDataRange().getDisplayValues(), h = v[0].map(function (x) { return String(x).trim(); });
  const col = function (n) { return h.indexOf(n); };
  const out = {};
  v.slice(1).forEach(function (r) {
    const sid = digits800_(r[col('Student 800#') > -1 ? col('Student 800#') : ML_800_INDEX_]);
    const day = ymdOf_(r[col('Consequence Date')] || '') || ymdOf_(r[col('Entry Date')] || '');
    const status = col('Status') > -1 ? String(r[col('Status')]).toLowerCase() : '';
    if (!sid || !day || /cancel|resched/.test(status)) return;
    (out[sid] = out[sid] || []).push({ day: day, type: String(r[col('Behavior Type')] || '').trim(), cons: String(r[col('Consequence Type')] || '') });
  });
  return out;
}

function daysBetween_(a, b) { return Math.round((new Date(b + 'T12:00:00') - new Date(a + 'T12:00:00')) / 86400000); }

// The school year a date is in starts on Aug 1.
function yearStartOf_(ymd) { const y = Number(ymd.slice(0, 4)), m = Number(ymd.slice(5, 7)); return (m >= 8 ? y : y - 1) + '-08-01'; }

function outcomeFor_(c, rows, today) {
  const panel = c.scheduled_date || String(c.closed_at || '').slice(0, 10);
  const start = yearStartOf_(panel);
  const before = rows.filter(function (r) { return r.day >= start && r.day < panel; });
  const after = rows.filter(function (r) { return r.day > panel && r.day <= today; });
  const monthsBefore = Math.max(1, daysBetween_(start, panel) / 30.4), monthsAfter = Math.max(0, daysBetween_(panel, today) / 30.4);
  const type = String(c.behavior_type || '').toLowerCase();
  const same = function (list) { return type ? list.filter(function (r) { return r.type.toLowerCase() === type; }).length : null; };
  const windows = {};
  OUTCOME_WINDOWS_.forEach(function (n) {
    if (daysBetween_(panel, today) < n) { windows[n] = null; return; } // not that long yet
    const pre = rows.filter(function (r) { return r.day >= ymdOf_(new Date(new Date(panel + 'T12:00:00').getTime() - n * 86400000)) && r.day < panel; });
    const post = rows.filter(function (r) { return r.day > panel && r.day <= ymdOf_(new Date(new Date(panel + 'T12:00:00').getTime() + n * 86400000)); });
    windows[n] = { before: pre.length, after: post.length, sameBefore: same(pre), sameAfter: same(post) };
  });
  let need = '';
  try { const s = c.panel_session_id ? panelRow_(c.panel_session_id) : null; need = s ? String(s.need_bucket || '') : ''; } catch (e) { need = ''; }
  return {
    case_id: c.case_id, student_name: c.student_name, student_id: c.student_id, grade: c.grade, squad: c.squad, behavior_type: c.behavior_type || '',
    need: need, panel: panel, status: c.status, status_label: c.status_label, summit: !!c.summit_at,
    perMonthBefore: Math.round(10 * before.length / monthsBefore) / 10,
    perMonthAfter: monthsAfter >= 0.5 ? Math.round(10 * after.length / monthsAfter) / 10 : null,
    after: after.length, detentionsAfter: after.filter(function (r) { return /detention/i.test(r.cons); }).length,
    sameAfter: same(after), windows: windows,
    daysToFinish: c.status === 'closed_completed' && c.closed_at ? Math.max(0, daysBetween_(panel, String(c.closed_at).slice(0, 10))) : null,
  };
}

function sumUp_(list) {
  const n = list.length, avg = function (f) { const v = list.map(f).filter(function (x) { return x != null; }); return v.length ? Math.round(10 * v.reduce(function (a, b) { return a + b; }, 0) / v.length) / 10 : null; };
  const windows = {};
  OUTCOME_WINDOWS_.forEach(function (w) {
    const have = list.filter(function (o) { return o.windows[w]; });
    windows[w] = have.length ? { students: have.length,
      noRepeat: Math.round(100 * have.filter(function (o) { return (o.windows[w].sameAfter || 0) === 0; }).length / have.length),
      before: have.reduce(function (a, o) { return a + o.windows[w].before; }, 0), after: have.reduce(function (a, o) { return a + o.windows[w].after; }, 0) } : null;
  });
  return { students: n, perMonthBefore: avg(function (o) { return o.perMonthBefore; }), perMonthAfter: avg(function (o) { return o.perMonthAfter; }),
    finished: list.filter(function (o) { return o.status === 'closed_completed'; }).length, windows: windows };
}

/** The dashboard's numbers. */
function trailbackOutcomes_(body) {
  if (!canSeeOutcomes_(body._staffEmail)) throw new Error('The outcomes are for the panel leads and the coordinator.');
  const today = todayYmd_(), cases = allCases_(), log = masterLogByStudent_();
  const count = function (f) { return cases.filter(f).length; };
  const heard = cases.filter(function (c) { return ['heard_approved', 'repair_in_progress', 'closed_completed', 'closed_reverted'].indexOf(c.status) > -1; });
  const outcomes = heard.filter(function (c) { return c.student_id; }).map(function (c) { return outcomeFor_(c, log[c.student_id] || [], today); });
  const group = function (key) {
    const g = {};
    outcomes.forEach(function (o) { const k = o[key] || 'Not recorded'; (g[k] = g[k] || []).push(o); });
    return Object.keys(g).sort().map(function (k) { return Object.assign({ label: k }, sumUp_(g[k])); });
  };
  let onTime = 0, logged = 0;
  cases.forEach(function (c) { Object.keys(c.checkin_notes || {}).forEach(function (n) { logged++; if (c.checkin_notes[n] && c.checkin_notes[n].on_time !== false) onTime++; }); });
  const finished = outcomes.filter(function (o) { return o.daysToFinish != null; });
  return {
    asOf: today,
    pipeline: { suggested: count(function (c) { return c.status === 'suggested'; }), offered: count(function (c) { return c.status === 'offered'; }),
      waiting: count(function (c) { return ['consented', 'tj_done', 'scheduled'].indexOf(c.status) > -1; }),
      inRepair: count(function (c) { return c.status === 'repair_in_progress' || c.status === 'heard_approved'; }),
      completed: count(function (c) { return c.status === 'closed_completed'; }), reverted: count(function (c) { return c.status === 'closed_reverted'; }),
      declined: count(function (c) { return c.status === 'declined'; }), summits: count(function (c) { return !!c.summit_at; }),
      detentionsCleared: count(function (c) { return c.status === 'closed_completed' && c.master_log_row; }) },
    checkinsOnTime: logged ? Math.round(100 * onTime / logged) : null,
    daysToFinish: finished.length ? Math.round(finished.reduce(function (a, o) { return a + o.daysToFinish; }, 0) / finished.length) : null,
    overall: sumUp_(outcomes), byBehavior: group('behavior_type'), byNeed: group('need'),
    students: outcomes.sort(function (a, b) { return a.panel < b.panel ? 1 : -1; }),
  };
}

function showOutcomesLink() {
  let url = '';
  try { url = ScriptApp.getService().getUrl() + '?page=outcomes'; } catch (e) { url = ''; }
  safeAlert_('Trailback outcomes (panel leads and the coordinator):\n\n' + url);
}
