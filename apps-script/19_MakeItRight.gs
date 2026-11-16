/**
 * Trail Journal -- Apps Script backend
 * File 19: making it right, after the panel.
 *
 * 1. LOGICAL CONSEQUENCE IDEAS IN THE PLAN BUILDER. The Accountability Log's "Logical Consequence
 *    Ideas" tab (staff add their own there) shows in the panel's plan builder next to the Action
 *    Ideas, matched to the case's behavior when the words line up (actionIdeas_ in 15_GiveGet.gs).
 *
 * 2. STUDENT SELF CHECK-IN (?page=selfcheck, public link). On the next-day and 10-day check-in
 *    dates the student gets an email with the link. They type their 800# and first name, see their
 *    plan actions, and say how each is going (done / partly / not yet + a sentence). It is saved on
 *    the case (student_checkins) and shows on the adult's one-tap check-in page. It never logs the
 *    official check-in by itself: the adult still does that.
 *
 * 3. "WAS THE HARM REPAIRED?" On the 10-day check-in date, the affected staff (case field
 *    affected_staff, else the referring staff's email) get one email with three one-click answers:
 *    Yes / Partly / Not yet (?page=repaired&t=...&a=...). Once per case. Students harmed are asked in
 *    person by the check-in adult, not by email.
 *
 * Emails only go out when TRAILBACK_EMAILS is ON (sendTrailbackMail_).
 */

// ── 1. IDEAS FROM THE ACCOUNTABILITY LOG ───────────────────────────
function logicalIdeasFromAccountability_() {
  const id = webSetting_('ACCOUNTABILITY_SHEET_ID');
  if (!id) return [];
  let sh;
  try { sh = SpreadsheetApp.openById(id).getSheetByName('Logical Consequence Ideas'); } catch (e) { return []; }
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getDataRange().getValues().slice(1).filter(function (r) { return String(r[1] || '').trim(); }).map(function (r) {
    return { menu: 'Repair', type: '', idea: String(r[1]).trim(), test: String(r[2] || '').trim() ? 'With: ' + String(r[2]).trim() : '',
      need: '', behavior: String(r[0] || '').trim(), source: 'Logical Consequence Ideas' };
  });
}

// ── 2. STUDENT SELF CHECK-IN ───────────────────────────────────────
function selfcheckUrl_() {
  const base = webSetting_('STUDENT_EXEC_URL') || (function () { try { return ScriptApp.getService().getUrl(); } catch (e) { return ''; } })();
  return base + '?page=selfcheck';
}

function firstWord_(s) { return String(s || '').trim().split(/[\s,]+/).filter(String)[0] || ''; }

// "Quill, Pat" -> Pat; "Pat Quill" -> Pat
function firstNameOf_(name) {
  const n = String(name || '').trim();
  return n.indexOf(',') > -1 ? firstWord_(n.split(',').slice(1).join(' ')) : firstWord_(n);
}

function selfcheckCase_(p) {
  const sid = digits800_(p && p.id);
  const first = String((p && p.first) || '').trim().toLowerCase();
  if (!sid || !first) throw new Error('Type your 800# and your first name.');
  const c = allCases_().filter(function (x) {
    return x.status === 'repair_in_progress' && digits800_(x.student_id) === sid && firstNameOf_(x.student_name).toLowerCase() === first;
  })[0];
  if (!c) throw new Error('No open Trailback plan for that 800# and first name. Check them, or ask your check-in adult.');
  return c;
}

function studentCheckinsOf_(c) {
  let a = c.student_checkins;
  if (typeof a === 'string') { try { a = JSON.parse(a || '[]'); } catch (e) { a = []; } }
  return Array.isArray(a) ? a : [];
}

/** Public: { id, first } -> the student's plan actions (no adult names or emails). */
function selfcheckLookup_(body) {
  const c = selfcheckCase_(body.payload || {});
  const last = studentCheckinsOf_(c).slice(-1)[0] || null;
  return { first: firstNameOf_(c.student_name), actions: repairActionsOf_(c).map(function (a) { return { id: a.id, action: a.action, due: a.due || '', test: a.test || '' }; }),
    last: last ? { at: last.at } : null };
}

/** Public: { id, first, answers: { a1: 'done'|'partly'|'not_yet' }, notes: { a1: '' }, note } */
function selfcheckSave_(body) {
  const p = body.payload || {}, c = selfcheckCase_(p);
  const ok = ['done', 'partly', 'not_yet'];
  const answers = {}, notes = {};
  repairActionsOf_(c).forEach(function (a) {
    const v = (p.answers || {})[a.id];
    if (ok.indexOf(v) > -1) answers[a.id] = v;
    const n = String((p.notes || {})[a.id] || '').trim().slice(0, 300);
    if (n) notes[a.id] = n;
  });
  if (!Object.keys(answers).length) throw new Error('Pick how each action is going.');
  const list = studentCheckinsOf_(c).concat([{ at: nowIso_(), answers: answers, notes: notes, note: String(p.note || '').trim().slice(0, 600) }]).slice(-10);
  genericUpdate_(TABS.PANEL_CASES, { case_id: c.case_id }, { student_checkins: JSON.stringify(list) });
  return { ok: true };
}

// ── 3. "WAS THE HARM REPAIRED?" ────────────────────────────────────
function affectedStaffEmails_(c) {
  const list = String(c.affected_staff || '').split(/[,;\s]+/).map(emailOf_).filter(String);
  if (!list.length && emailOf_(c.referring_staff)) list.push(emailOf_(c.referring_staff));
  return list.filter(function (e, i) { return list.indexOf(e) === i; }).join(',');
}

function repairedUrl_(token, answer) {
  const base = webSetting_('STUDENT_EXEC_URL') || (function () { try { return ScriptApp.getService().getUrl(); } catch (e) { return ''; } })();
  return base + '?page=repaired&t=' + encodeURIComponent(token) + '&a=' + answer;
}

function harmEmailHtml_(c, token) {
  const t = squadTheme_(c.squad), first = firstNameOf_(c.student_name) || 'The student';
  const btn = function (a, label, bg) {
    return '<a href="' + escHtml_(repairedUrl_(token, a)) + '" style="display:inline-block;margin:4px 6px 4px 0;padding:10px 18px;border-radius:999px;background:' + bg +
      ';color:#fff;font-family:Verdana,sans-serif;font-weight:bold;text-decoration:none">' + label + '</a>';
  };
  const actions = repairActionsOf_(c).map(function (a) { return emailInfoRow_('Action', escHtml_(a.action), t); }).join('');
  return emailShell_('&#129309;', 'Was the harm repaired?', escHtml_(first) + ', 10 school days after the Trailback Panel',
    '<p style="font-family:Verdana,sans-serif;font-size:14px;color:#3a3a3a;line-height:1.6">You were affected by what happened. ' + escHtml_(first) +
    ' has been working on a plan to make it right. One click tells us how it feels from your side:</p>' + actions +
    '<p>' + btn('yes', 'Yes, it is repaired', '#2e7d32') + btn('partly', 'Partly', '#b8821e') + btn('no', 'Not yet', '#9e2a2a') + '</p>' +
    emailCallout_('Private', 'Only the Trailback coordinator sees your answer. It does not go to the student or family.', null, t), t);
}

/** Daily (trailbackDailyCheck): the 10-day email, once per case. preview = build, send nothing. */
function harmRepairedEmails_(preview) {
  const today = todayYmd_(), out = [];
  allCases_().filter(function (c) { return c.status === 'repair_in_progress' && c.checkin_10 && ymdOf_(c.checkin_10) <= today && !c.harm_asked_at; }).forEach(function (c) {
    const to = affectedStaffEmails_(c);
    if (!to) return;
    const token = c.harm_token || Utilities.getUuid().replace(/-/g, '');
    const r = sendTrailbackMail_(to, 'Trailback: was the harm repaired? (' + (firstNameOf_(c.student_name) || c.case_id) + ')', harmEmailHtml_(c, token), preview);
    if (r.sent) genericUpdate_(TABS.PANEL_CASES, { case_id: c.case_id }, { harm_token: token, harm_asked_at: nowIso_() });
    out.push(r);
  });
  return out;
}

/** Public, by the one-time link: { t, a: yes|partly|no, note }. A later click changes the answer. */
function harmAnswer_(body) {
  const p = body.payload || {}, t = String(p.t || '').trim();
  if (t.length < 16) throw new Error('This link is not complete.');
  const r = genericQuery_(TABS.PANEL_CASES, { harm_token: t }, 1)[0];
  if (!r) throw new Error('This link has expired. Thank you anyway!');
  const c = caseOut_(r);
  const a = ['yes', 'partly', 'no'].indexOf(p.a) > -1 ? p.a : '';
  if (!a) throw new Error('Pick Yes, Partly or Not yet.');
  const patch = { harm_repaired: a, harm_answered_at: nowIso_() };
  if (String(p.note || '').trim()) patch.harm_note = String(p.note).trim().slice(0, 600);
  genericUpdate_(TABS.PANEL_CASES, { case_id: c.case_id }, patch);
  return { ok: true, first: firstNameOf_(c.student_name), answer: a };
}

/** Menu: preview the 10-day emails (nothing sent). */
function previewHarmEmails() {
  const r = harmRepairedEmails_(true);
  previewDialog_(r[0] || { skipped: 'No case is at its 10-day check-in with an affected staff email.' }, 'Was the harm repaired? (preview)');
}

// The student's self check-in link, the morning of the next-day and 10-day check-ins.
function selfcheckReminders_(preview) {
  const today = todayYmd_(), props = PropertiesService.getScriptProperties(), out = [];
  allCases_().filter(function (c) { return c.status === 'repair_in_progress'; }).forEach(function (c) {
    [3, 10].forEach(function (n) {
      const key = 'tb-self-' + c.case_id + '-' + n;
      if (ymdOf_(c['checkin_' + n]) !== today || (!preview && props.getProperty(key))) return;
      const to = studentEmail_(c.student_id);
      if (!to) return;
      const t = squadTheme_(c.squad), first = firstNameOf_(c.student_name) || 'there';
      const r = sendTrailbackMail_(to, 'Your Trailback check-in', emailShell_('&#9989;', 'Your Trailback check-in', 'Hi ' + escHtml_(first),
        emailCallout_('How is your plan going?', '<a href="' + escHtml_(selfcheckUrl_()) + '">Open your check-in</a>, type your 800# and first name, and mark each action. Your check-in adult will see it.', null, t), t), preview);
      if (r.sent) props.setProperty(key, today);
      out.push(r);
    });
  });
  return out;
}
