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
 * REDEMPTION, the moment the Summit is recorded:
 *   - TRAILBACK_REDEMPTION_POINTS (75) go to the student's squad on Clash right away, as a
 *     'redemption' row: "a Gold climber" unless the student said yes to their first name and last initial.
 *   - A private row on Clash's Redemptions tab gives the student a Summit Climber badge and LRG
 *     points on My Climb (Clash 04_Students.gs / 19_Evolve.gs).
 *   - The family gets the Summit card by email (when the case has a family email).
 *   - TRAILBACK_GUIDE_DAYS (60) school days later with no new Master Log entry, the student is
 *     eligible to be a Trail Guide (a student panelist); the coordinator is told (17_Pipeline.gs).
 */

function redemptionPoints_() { const t = String(webSetting_('TRAILBACK_REDEMPTION_POINTS')).trim(); return t === '' ? 75 : Math.max(0, Number(t) || 0); }
function titleSquad2_(s) { s = String(s || '').toUpperCase(); return s.charAt(0) + s.slice(1).toLowerCase(); }

/** The squad's Redemption points and the student's private Redemptions row on Clash. Returns warnings. */
function postRedemption_(c, shareName) {
  const pts = redemptionPoints_(), squad = String(c.squad || '').trim().toUpperCase();
  if (!pts) return [];
  if (c.redemption_at) return [];
  if (COC_SQUADS_.indexOf(squad) === -1) return ['No squad on this case, so no Redemption points went to Clash.'];
  try {
    const book = SpreadsheetApp.openById(COC_SHEET_ID_), ev = book.getSheetByName(COC_EVENTS_TAB_);
    if (!ev) return ['The Clash sheet has no events tab.'];
    const today = Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd');
    const p = String(c.student_name || '').trim().split(/\s+/);
    const who = shareName && p[0] ? p[0] + (p.length > 1 ? ' ' + p[p.length - 1].charAt(0) + '.' : '') + ' (' + titleSquad2_(squad) + ')' : 'a ' + titleSquad2_(squad) + ' climber';
    ev.appendRow(['', 'Redemption: ' + who + ' reached the Summit', 'redemption', today, 'Trailback Summit', '',
      squad === 'BLACK' ? pts : '', squad === 'GOLD' ? pts : '', squad === 'GREY' ? pts : '', squad === 'PURPLE' ? pts : '']);
    let red = book.getSheetByName('Redemptions');
    if (!red) { red = book.insertSheet('Redemptions'); red.appendRow(['800#', 'Squad', 'Date', 'Points', 'Case']); }
    if (c.student_id) red.appendRow([String(c.student_id), squad, today, pts, c.case_id]);
    genericUpdate_(TABS.PANEL_CASES, { case_id: c.case_id }, { redemption_at: nowIso_(), redemption_points: pts });
    return [];
  } catch (e) { return ['Redemption points not posted: ' + e.message]; }
}

/** The Summit card to the family. Returns warnings. */
function familySummitEmail_(c, s) {
  const to = emailOf_(c.family_email);
  if (!to || c.family_summit_sent_at) return [];
  try {
    const t = squadTheme_(c.squad), first = String(c.student_name || '').trim().split(/\s+/)[0] || 'Your student';
    const row = function (label, text) { return text ? emailInfoRow_(label, escHtml_(text), t) : ''; };
    const html = emailShell_('&#127956;', first + ' reached the Summit', 'A Trailback Panel plan, finished',
      '<p style="font-family:Verdana,sans-serif;font-size:14px;color:#3a3a3a;line-height:1.6">We wanted you to hear this from us: ' + escHtml_(first) +
      ' finished every part of the plan the Trailback Panel made, and came back to tell us about it. The detention is cleared. We are proud of this, and we hope you are too.</p>' +
      repairActionsOf_(c).map(function (a) { return emailInfoRow_('✓', escHtml_(a.action), t); }).join('') +
      row('What ' + first + ' accomplished', s.accomplished) + row('What ' + first + ' learned', s.learned) + row('What ' + first + ' is proud of', s.proud) +
      row('Who helped', s.helped) + row('What is next', s.next) + row('What the adults noticed', s.noticed) +
      emailCallout_('Say it at home too', 'Ask ' + escHtml_(first) + ' about it tonight. Hearing it from you matters.', null, t), t);
    const r = sendTrailbackMail_(to, first + ' reached the Summit', html, false);
    if (r.sent) genericUpdate_(TABS.PANEL_CASES, { case_id: c.case_id }, { family_summit_sent_at: nowIso_() });
    return r.sent ? [] : quietOff_(['Summit card not emailed home: ' + r.why]);
  } catch (e) { return ['Summit card not emailed home: ' + e.message]; }
}

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
  return { case_id: c.case_id, student_name: c.student_name, grade: c.grade, squad: c.squad, closed_at: String(c.closed_at || '').slice(0, 10), points: redemptionPoints_(), family: !!emailOf_(c.family_email),
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
  s.shareName = p.share_name === true || p.share_name === 'yes';
  if (!s.accomplished && !s.proud) throw new Error('Write what they accomplished or what they are proud of.');
  saveCase_(c, { summit_at: nowIso_(), summit_json: JSON.stringify(s) }, me, c.status, 'Summit celebration held');
  const fresh = getCase_(c.case_id);
  const warnings = postRedemption_(fresh, s.shareName).concat(familySummitEmail_(fresh, s));
  return Object.assign(summitView_(getCase_(c.case_id)), { redemption: redemptionPoints_(), warnings: warnings });
}
