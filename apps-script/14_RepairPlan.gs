/**
 * Trail Journal -- Apps Script backend
 * File 14: the panel's questions, the repair plan tracker, one-tap check-ins and the reminders.
 *
 * PANEL QUESTIONS tab (staff edit any time): the v2 question sheet, one row per question.
 *   Show | Stage | Behavior | When | Question
 *   Stage: Strengths, Story, Impact, Themes, Making it right, Back pocket
 *   Behavior: blank = every case, or a Behavior type from the consequence form (Disrespect,
 *     Disruption, Hallway / lunch conduct, Technology, Physical contact). The panel page shows the
 *     everyday questions plus the ones for the case's behavior type.
 *   When: Ask first, If there is time, After the impact statement
 * THE PLAN (made on the panel's Lock the plan screen, approved by the AP): one to three actions,
 *   each with how we know it is done, a due date (10 school days by default) and a check-in adult
 *   (an email). At submit the actions go onto the case (panel_cases.repair_actions):
 *   [{ id, menu, action, test, due, adult, status: open|done|partly|not_yet, note, updated_at, by }]
 * ONE-TAP CHECK-IN (?page=checkin on the staff link): the check-in adult (or a panel lead) sees
 *   their open cases, marks each action Done / Partly / Not yet, picks On track / Needs help /
 *   Not done for the check-in that is due, and saves. That logs the 3/10/30 check-in on the case.
 *   Every action Done = the case closes as completed by itself (TRAILBACK_AUTO_CLOSE, ON): the
 *   Master Log row becomes Resolved (Trailback) and the anonymous Clash point is requested, the
 *   same as closing it by hand. Nothing ever reverts by itself: that stays a person's call.
 * REMINDERS (the weekday 7 a.m. check, when TRAILBACK_EMAILS is ON):
 *   - the morning a check-in is due, the check-in adult gets the one-tap link;
 *   - two school days before an action is due and it isn't done: the adult and the coordinator;
 *   - an action past due and not done: the coordinator, to extend it or revert the case.
 */

const PANEL_QUESTIONS_TAB_ = 'Panel Questions';
const PANEL_STAGES_ = ['Strengths', 'Story', 'Impact', 'Themes', 'Making it right', 'Back pocket'];
const ACTION_STATUSES_ = ['open', 'done', 'partly', 'not_yet'];

// [Stage, Behavior, When, Question]. The v2 question sheet, then a few starters per behavior type.
const PANEL_QUESTIONS_SEED_ = [
  ['Strengths', '', 'Ask first', 'What are you good at that people in this room might not know?'],
  ['Strengths', '', 'Ask first', 'Where in this building do you feel like you belong?'],
  ['Strengths', '', 'If there is time', 'When are you at your best at school? Who sees it?'],
  ['Strengths', '', 'If there is time', 'What is one thing you are really proud of?'],
  ['Strengths', '', 'If there is time', 'What is the last thing you worked really hard on?'],
  ['Strengths', '', 'If there is time', 'Who do you get along with best here, and why?'],
  ['Strengths', '', 'If there is time', 'What do you like doing outside school?'],
  ['Story', '', 'Ask first', 'What happened, from your side?'],
  ['Story', '', 'Ask first', 'What were you thinking at the time?'],
  ['Story', '', 'If there is time', 'What have you thought about since?'],
  ['Story', '', 'If there is time', 'What role do you think you played in this?'],
  ['Story', '', 'If there is time', 'What do you wish the adults at school knew about this?'],
  ['Impact', '', 'Ask first', 'Who do you think you impacted? Start with the person closest to it.'],
  ['Impact', '', 'Ask first', 'Who else inside school? Classmates, your teacher, other staff, yourself.'],
  ['Impact', '', 'Ask first', 'Who outside this building? Family at home, your team or club, neighbors, people online.'],
  ['Impact', '', 'Ask first', 'How was each of them affected?'],
  ['Impact', '', 'If there is time', 'What was it like to be that person in that moment?'],
  ['Impact', '', 'If there is time', 'What has this cost you?'],
  ['Impact', '', 'If there is time', 'Whose trust do you most want to earn back?'],
  ['Impact', '', 'If there is time', 'How has this followed you home or online?'],
  ['Impact', '', 'If there is time', 'What do you want people outside school to see in how you handle this?'],
  ['Impact', '', 'If there is time', 'Is there anyone you had not thought about until now?'],
  ['Impact', '', 'After the impact statement', 'What stood out in what you just heard?'],
  ['Impact', '', 'After the impact statement', 'What do you want that person to know?'],
  ['Themes', '', 'Ask first', 'Has something like this come up before this year?'],
  ['Themes', '', 'Ask first', 'What do those times have in common? Class, time of day, who you were with, how you were feeling.'],
  ['Themes', '', 'If there is time', 'What was going on that led up to it this time?'],
  ['Themes', '', 'If there is time', 'What have adults tried? What helped, even a little?'],
  ['Themes', '', 'If there is time', 'When has it gone well? What was different?'],
  ['Themes', '', 'If there is time', 'What would you call this pattern?'],
  ['Themes', '', 'If there is time', 'What is one thing about your day that would make this less likely?'],
  ['Making it right', '', 'Ask first', 'What would make this right?'],
  ['Making it right', '', 'Ask first', 'What do you need so this doesn\'t happen again?'],
  ['Making it right', '', 'If there is time', 'What end result would you like to see?'],
  ['Making it right', '', 'If there is time', 'What do you want to be true in 30 school days?'],
  ['Making it right', '', 'If there is time', 'What are two ways you could get there?'],
  ['Making it right', '', 'If there is time', 'From 1 to 10, how confident are you? What would move it up one?'],
  ['Making it right', '', 'If there is time', 'Which adult here do you trust to check in with you?'],
  ['Back pocket', '', 'Shift the view', 'How would you feel if this happened to you?'],
  ['Back pocket', '', 'Shift the view', 'How might someone else respond to that?'],
  ['Back pocket', '', 'Make it familiar', 'What advice would you give a friend who did this?'],
  ['Back pocket', '', 'Make it familiar', 'Tell me about a time this went differently.'],
  ['Back pocket', '', 'Clear it up', 'Can you explain that a different way?'],
  ['Back pocket', '', 'Clear it up', 'How has your thinking changed since it happened?'],
  // By behavior type (starters: edit, add or untick Show on the tab).
  ['Impact', 'Disrespect', 'Ask first', 'When you said or did that, what did you want the other person to understand?'],
  ['Impact', 'Disrespect', 'If there is time', 'How do you think the rest of the class saw that moment?'],
  ['Themes', 'Disrespect', 'If there is time', 'Which adults do you get along with? What do they do differently?'],
  ['Impact', 'Disruption', 'Ask first', 'What did the class lose while it was happening?'],
  ['Themes', 'Disruption', 'If there is time', 'What usually happens right before you get pulled off task?'],
  ['Making it right', 'Disruption', 'If there is time', 'What signal could your teacher give you that would actually help?'],
  ['Impact', 'Hallway / lunch conduct', 'Ask first', 'Who else was in the hallway or at the table, and how did it land for them?'],
  ['Themes', 'Hallway / lunch conduct', 'If there is time', 'Who are you usually with when this happens?'],
  ['Making it right', 'Hallway / lunch conduct', 'If there is time', 'What would make lunch or the hallway go better for you?'],
  ['Impact', 'Technology', 'Ask first', 'Who saw it online, and who could still see it now?'],
  ['Impact', 'Technology', 'If there is time', 'How is a screen different from saying it face to face?'],
  ['Making it right', 'Technology', 'If there is time', 'What would you change about how you use your phone or Chromebook at school?'],
  ['Impact', 'Physical contact', 'Ask first', 'How did the other person feel in their body when it happened?'],
  ['Impact', 'Physical contact', 'If there is time', 'Who else felt less safe because of it?'],
  ['Themes', 'Physical contact', 'If there is time', 'What were you feeling right before? Where do you feel that in your body?'],
];

function ensurePanelQuestionsTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(PANEL_QUESTIONS_TAB_)) return;
  const sh = ss.insertSheet(PANEL_QUESTIONS_TAB_);
  const head = ['Show', 'Stage', 'Behavior', 'When', 'Question'];
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#2c4a35').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.getRange(2, 1, PANEL_QUESTIONS_SEED_.length, head.length).setValues(PANEL_QUESTIONS_SEED_.map(function (r) { return [true].concat(r); }));
}

/** The questions for a case: everyday ones plus the case's behavior type, grouped by stage. */
function panelQuestions_(body) {
  if (!body._staffEmail && !body._admin) throw new Error('Staff only.');
  ensurePanelQuestionsTab_();
  const behavior = String((body.payload || {}).behavior || '').trim().toLowerCase();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PANEL_QUESTIONS_TAB_);
  const rows = sh.getDataRange().getValues();
  const out = {};
  PANEL_STAGES_.forEach(function (s) { out[s] = []; });
  rows.slice(1).forEach(function (r) {
    const show = r[0] === true || /^(true|yes|x)$/i.test(String(r[0]));
    const stage = String(r[1] || '').trim(), b = String(r[2] || '').trim(), q = String(r[4] || '').trim();
    if (!show || !q || !out[stage]) return;
    if (b && (!behavior || behavior.indexOf(b.toLowerCase()) === -1 && b.toLowerCase().indexOf(behavior) === -1)) return;
    out[stage].push({ when: String(r[3] || '').trim() || 'If there is time', q: q, behavior: b });
  });
  return { stages: out, behavior: behavior };
}

// ── THE PLAN ON THE CASE ────────────────────────────────────────────
function emailOf_(s) { const m = String(s || '').match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/); return m ? m[0].toLowerCase() : ''; }

/** The plan actions from a panel record, cleaned up for the case (1 to 3, each with a due date). */
function planActionsFrom_(record, by) {
  let list = record.plan_actions;
  if (typeof list === 'string') { try { list = JSON.parse(list); } catch (e) { list = []; } }
  if (!Array.isArray(list)) list = [];
  const fallbackDue = ymdOf_((record.followup_dates || {}).checkin_10) || '';
  return list.filter(function (a) { return a && String(a.action || '').trim(); }).slice(0, 3).map(function (a, i) {
    return { id: 'a' + (i + 1), menu: String(a.menu || ''), type: String(a.type || ''), action: String(a.action).trim().slice(0, 300),
      test: String(a.test || '').trim().slice(0, 300), due: ymdOf_(a.due) || fallbackDue, adult: String(a.adult || '').trim().slice(0, 120),
      status: 'open', note: '', updated_at: nowIso_(), by: by || '' };
  });
}

function repairActionsOf_(c) {
  let a = c.repair_actions;
  if (typeof a === 'string') { try { a = JSON.parse(a || '[]'); } catch (e) { a = []; } }
  return Array.isArray(a) ? a : [];
}

// Everyone who checks in on a case: the check-in person and each action's adult (emails only).
function caseAdults_(c) {
  const out = [];
  [c.checkin_person].concat(repairActionsOf_(c).map(function (a) { return a.adult; })).forEach(function (s) {
    const e = emailOf_(s);
    if (e && out.indexOf(e) === -1) out.push(e);
  });
  return out;
}

// The check-in that is due now: the earliest of 3/10/30 whose date has come and isn't logged
// (or the next one coming, so a check-in can be logged early).
function dueCheckin_(c) {
  const today = todayYmd_();
  const open = CHECKIN_SCHOOL_DAYS_.filter(function (n) { return !c.checkin_notes[n]; });
  const due = open.filter(function (n) { return c['checkin_' + n] && c['checkin_' + n] <= today; });
  return due.length ? due[0] : (open[0] || null);
}

function checkinView_(c) {
  const n = dueCheckin_(c);
  return { case_id: c.case_id, student_name: c.student_name, grade: c.grade, squad: c.squad, behavior_type: c.behavior_type,
    actions: repairActionsOf_(c), checkin_person: c.checkin_person, completion_criteria: c.completion_criteria,
    next_checkin: n ? { n: n, date: c['checkin_' + n] || '' } : null, logged: Object.keys(c.checkin_notes || {}),
    checkins: { 3: c.checkin_3, 10: c.checkin_10, 30: c.checkin_30 } };
}

/** The check-in page: the signed-in adult's open cases (panel leads and the coordinator see all). */
function repairCheckins_(body) {
  const me = String(body._staffEmail || '').toLowerCase();
  if (!me) throw new Error('Open the check-in page from the staff link and sign in with your school Google account.');
  const seesAll = panelIsAdmin_(me) || coordinatorEmails_().toLowerCase().split(',').indexOf(me) > -1;
  const mine = allCases_().filter(function (c) { return c.status === 'repair_in_progress' && (seesAll || caseAdults_(c).indexOf(me) > -1); });
  return { me: me, seesAll: seesAll, cases: mine.map(checkinView_) };
}

/**
 * One tap: { case_id, actions: { a1: 'done'|'partly'|'not_yet', ... }, notes: { a1: '' }, overall:
 * 'on_track'|'needs_help'|'not_done', note }. Logs the due check-in; every action done = closed.
 */
function repairCheckin_(body) {
  const me = String(body._staffEmail || '').toLowerCase();
  if (!me) throw new Error('Sign in with your school Google account.');
  const p = body.payload || {};
  const c = getCase_(p.case_id);
  if (c.status !== 'repair_in_progress') throw new Error('This case is ' + c.status_label + '.');
  const seesAll = panelIsAdmin_(me) || coordinatorEmails_().toLowerCase().split(',').indexOf(me) > -1;
  if (!seesAll && caseAdults_(c).indexOf(me) === -1) throw new Error('You are not a check-in adult on this case.');
  const marks = p.actions || {}, notes = p.notes || {};
  const actions = repairActionsOf_(c).map(function (a) {
    const s = String(marks[a.id] || '');
    if (ACTION_STATUSES_.indexOf(s) === -1 || s === 'open') return a;
    return Object.assign({}, a, { status: s, note: String(notes[a.id] == null ? a.note : notes[a.id]).slice(0, 300), updated_at: nowIso_(), by: me });
  });
  const overall = { on_track: 'On track', needs_help: 'Needs help', not_done: 'Not done' }[p.overall] || '';
  genericUpdate_(TABS.PANEL_CASES, { case_id: c.case_id }, { repair_actions: JSON.stringify(actions), updated_by: me, updated_at: nowIso_() });
  const out = { logged: null, closed: false, warnings: [] };
  const n = dueCheckin_(c);
  const summary = actions.map(function (a) { return a.action + ': ' + ({ done: 'done', partly: 'partly', not_yet: 'not yet', open: 'open' })[a.status]; }).join('; ');
  if (n) {
    updateTrailbackCase_({ _staffEmail: me, payload: { case_id: c.case_id, op: 'checkin', n: n, note: [overall, summary, String(p.note || '').trim()].filter(String).join(' · ') } });
    out.logged = n;
  }
  const allDone = actions.length && actions.every(function (a) { return a.status === 'done'; });
  if (allDone && String(webSetting_('TRAILBACK_AUTO_CLOSE') || 'ON').toUpperCase() !== 'OFF') {
    const r = updateTrailbackCase_({ _staffEmail: me, payload: { case_id: c.case_id, op: 'close_completed', reason: 'Every action done (checked in by ' + me + ')' } });
    out.closed = true; out.warnings = r.warnings || [];
  } else if (overall === 'Needs help' || overall === 'Not done') {
    const to = coordinatorEmails_();
    if (to) sendTrailbackMail_(to, 'Trailback check-in: ' + (c.student_name || c.case_id) + ' ' + overall.toLowerCase(),
      emailShell_('&#9888;', 'A check-in needs you', escHtml_(c.student_name || c.case_id), emailInfoRow_('Check-in', overall + ' (by ' + escHtml_(me) + ')') +
        emailInfoRow_('Actions', escHtml_(summary)) + (p.note ? emailInfoRow_('Note', escHtml_(p.note)) : ''), squadTheme_(c.squad)), false);
  }
  return Object.assign(out, { case: checkinView_(getCase_(c.case_id)) });
}

// ── REMINDERS (from the weekday 7 a.m. check) ───────────────────────
function checkinUrl_(caseId) {
  let base = '';
  try { base = ScriptApp.getService().getUrl(); } catch (e) { base = ''; }
  if (!base) base = String(webSetting_('SITE_URL') || '').replace(/\/+$/, '') + '/panel/';
  return base + '?page=checkin' + (caseId ? '&case=' + encodeURIComponent(caseId) : '');
}

function schoolDaysBetween_(from, to) {
  const off = noSchoolDates_();
  let n = 0;
  const d = new Date(from + 'T12:00:00');
  while (ymdOf_(d) < to && n < 400) {
    d.setDate(d.getDate() + 1);
    if (d.getDay() !== 0 && d.getDay() !== 6 && !off[ymdOf_(d)]) n++;
  }
  return n;
}

/** Returns what it sent (preview = build only). Each reminder goes once (Script Properties). */
function repairReminders_(preview) {
  const today = todayYmd_(), props = PropertiesService.getScriptProperties(), sent = [];
  const once = function (key) { if (preview) return true; if (props.getProperty(key)) return false; props.setProperty(key, today); return true; };
  allCases_().filter(function (c) { return c.status === 'repair_in_progress'; }).forEach(function (c) {
    const t = squadTheme_(c.squad), who = escHtml_(c.student_name || c.case_id), adults = caseAdults_(c).join(',');
    const actions = repairActionsOf_(c);
    // The morning a check-in is due.
    CHECKIN_SCHOOL_DAYS_.forEach(function (n) {
      if (c['checkin_' + n] !== today || c.checkin_notes[n] || !adults || !once('tb-ci-' + c.case_id + '-' + n)) return;
      sent.push(sendTrailbackMail_(adults, 'Trailback check-in today: ' + (c.student_name || c.case_id) + ' (' + n + ' school days)',
        emailShell_('&#9989;', n + '-school-day check-in', who,
          actions.map(function (a) { return emailInfoRow_(escHtml_(a.action), ({ done: '✓ done', partly: 'partly', not_yet: 'not yet', open: 'open' })[a.status] + (a.due ? ' · due ' + a.due : ''), t); }).join('') +
          emailCallout_('One tap', '<a href="' + escHtml_(checkinUrl_(c.case_id)) + '">Open the check-in</a>: mark each action and how it\'s going. Every action done closes the case.', null, t), t), preview));
    });
    actions.filter(function (a) { return a.status !== 'done' && a.due; }).forEach(function (a) {
      const left = a.due >= today ? schoolDaysBetween_(today, a.due) : -1;
      if (left === 2 && once('tb-due2-' + c.case_id + '-' + a.id)) {
        sent.push(sendTrailbackMail_([emailOf_(a.adult), coordinatorEmails_()].filter(String).join(','), 'Trailback: due in 2 school days, ' + (c.student_name || c.case_id),
          emailShell_('&#9203;', 'Due in 2 school days', who, emailInfoRow_('Action', escHtml_(a.action), t) + emailInfoRow_('How we know it is done', escHtml_(a.test || '—'), t) +
            emailInfoRow_('Due', a.due, t) + emailCallout_('Check in', '<a href="' + escHtml_(checkinUrl_(c.case_id)) + '">Mark it</a> when it is done.', null, t), t), preview));
      }
      if (a.due < today && once('tb-late-' + c.case_id + '-' + a.id)) {
        sent.push(sendTrailbackMail_(coordinatorEmails_(), 'Trailback: past due, ' + (c.student_name || c.case_id),
          emailShell_('&#9888;', 'An action is past due', who, emailInfoRow_('Action', escHtml_(a.action), t) + emailInfoRow_('Was due', a.due, t) +
            emailCallout_('Your call', 'Give more time, or close the case as reverted on Admin › Trailback Cases. Nothing reverts by itself.', null, t), t), preview));
      }
    });
  });
  return sent;
}


function showCheckinLink() {
  safeAlert_('The one-tap Trailback check-in page (staff sign-in):\n\n' + checkinUrl_('') + '\n\nCheck-in adults get this link by email the morning each check-in is due.');
}
