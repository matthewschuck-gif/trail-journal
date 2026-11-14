/**
 * Trail Journal -- Apps Script backend
 * File 16: the Summit celebration, when a student finishes every part of a Trailback plan.
 *
 * The student comes back to the panel (or a few of its members and the check-in adult) and
 * shares, humbly, what they did: what they accomplished, what they learned, what they are proud
 * of, who helped, and what is next. The adults name what they noticed.
 *
 *   - When a case closes as completed (by hand or by the last check-in), the case's panelists
 *     (sign-ups), check-in adults and the coordinator get an invitation: the questions to ask,
 *     the plan the student finished, and the link to record it. TRAILBACK_EMAILS must be ON.
 *   - The check-in page lists the Summits still to hold. Recording one saves it on the case
 *     (summit_at, summit_json) and gives a printable Summit card for the student to keep.
 */

const SUMMIT_QUESTIONS_ = [
  ['accomplished', 'What did you get done? Walk us through it.'],
  ['learned', 'What did you learn about yourself, or about the people it affected?'],
  ['proud', 'What are you most proud of?'],
  ['helped', 'Who helped you get here?'],
  ['next', 'What will you keep doing? What is next for you?'],
];

function summitRecipients_(c) {
  const to = caseAdults_(c);
  try {
    panelConfirmed_(panelSignupTable_().rows, c.case_id).forEach(function (s) { const e = String(s.staff_email || '').toLowerCase(); if (e && to.indexOf(e) === -1) to.push(e); });
  } catch (e) { /* no sign-ups tab yet */ }
  coordinatorEmails_().toLowerCase().split(',').forEach(function (e) { e = e.trim(); if (e && to.indexOf(e) === -1) to.push(e); });
  return to;
}

/** The invitation when a plan is finished. Returns warnings (nothing throws). */
function summitInvite_(c) {
  try {
    const t = squadTheme_(c.squad), who = escHtml_(c.student_name || c.case_id);
    const actions = repairActionsOf_(c);
    const html = emailShell_('&#127956;', 'Summit celebration', who + ' finished the Trailback plan',
      emailCallout_('Bring them back to celebrate', 'Invite ' + who + ' back to the panel (a few of you and the check-in adult is plenty) for a short Summit: ' +
        'the student shares what they did, humbly, and the adults name what they noticed. Ten minutes.', null, t) +
      (actions.length ? '<p style="font-family:Verdana,sans-serif;font-size:13px;color:#3a3a3a;margin:12px 0 4px"><strong>What they finished</strong></p>' +
        actions.map(function (a) { return emailInfoRow_('✓', escHtml_(a.action), t); }).join('') : '') +
      '<p style="font-family:Verdana,sans-serif;font-size:13px;color:#3a3a3a;margin:12px 0 4px"><strong>Ask</strong></p>' +
      SUMMIT_QUESTIONS_.map(function (q) { return emailInfoRow_('', escHtml_(q[1]), t); }).join('') +
      emailCallout_('Record it', '<a href="' + escHtml_(checkinUrl_(c.case_id)) + '">Open the check-in page</a> › Summit celebrations. You can print a Summit card for the student to keep.', null, t), t);
    const r = sendTrailbackMail_(summitRecipients_(c).join(','), 'Summit celebration: ' + (c.student_name || c.case_id) + ' finished the Trailback plan', html, false);
    return r.sent ? [] : ['Summit invitation not sent: ' + r.why];
  } catch (e) { return ['Summit invitation not sent: ' + e.message]; }
}

function summitView_(c) {
  let s = null;
  try { s = c.summit_json ? JSON.parse(c.summit_json) : null; } catch (e) { s = null; }
  return { case_id: c.case_id, student_name: c.student_name, grade: c.grade, squad: c.squad, closed_at: String(c.closed_at || '').slice(0, 10),
    actions: repairActionsOf_(c).map(function (a) { return a.action; }), summit_at: String(c.summit_at || ''), summit: s };
}

/** The Summits still to hold (and the last few held) that this adult can see. */
function summitList_(body) {
  const me = String(body._staffEmail || '').toLowerCase();
  if (!me) throw new Error('Sign in with your school Google account.');
  const seesAll = panelIsAdmin_(me) || coordinatorEmails_().toLowerCase().split(',').indexOf(me) > -1;
  const mine = allCases_().filter(function (c) { return c.status === 'closed_completed' && (seesAll || summitRecipients_(c).indexOf(me) > -1); });
  return {
    toHold: mine.filter(function (c) { return !c.summit_at; }).map(summitView_),
    held: mine.filter(function (c) { return c.summit_at; }).sort(function (a, b) { return String(b.summit_at).localeCompare(String(a.summit_at)); }).slice(0, 10).map(summitView_),
    questions: SUMMIT_QUESTIONS_.map(function (q) { return { key: q[0], q: q[1] }; }),
  };
}

/** { case_id, answers: { accomplished, learned, proud, helped, next }, noticed, present } */
function summitSave_(body) {
  const me = String(body._staffEmail || '').toLowerCase();
  if (!me) throw new Error('Sign in with your school Google account.');
  const p = body.payload || {};
  const c = getCase_(p.case_id);
  if (c.status !== 'closed_completed') throw new Error('A Summit is for a finished plan.');
  const seesAll = panelIsAdmin_(me) || coordinatorEmails_().toLowerCase().split(',').indexOf(me) > -1;
  if (!seesAll && summitRecipients_(c).indexOf(me) === -1) throw new Error('You are not on this case.');
  const a = p.answers || {}, s = {};
  SUMMIT_QUESTIONS_.forEach(function (q) { s[q[0]] = String(a[q[0]] || '').trim().slice(0, 800); });
  s.noticed = String(p.noticed || '').trim().slice(0, 800);
  s.present = String(p.present || '').trim().slice(0, 300);
  s.by = me;
  if (!s.accomplished && !s.proud) throw new Error('Write what they accomplished or what they are proud of.');
  saveCase_(c, { summit_at: nowIso_(), summit_json: JSON.stringify(s) }, me, c.status, 'Summit celebration held');
  return summitView_(getCase_(c.case_id));
}
