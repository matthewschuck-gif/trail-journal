/**
 * Trail Journal -- Apps Script backend
 * File 17: the Trailback pipeline on autopilot, from spotting a student to a full panel.
 *
 * Every weekday morning (the 7 a.m. check, TRAILBACK_AUTOMATION ON):
 *   1. SPOT  - Repeat pattern: a student with TRAILBACK_PATTERN_COUNT (3) or more Master Log rows
 *              with the same Behavior Type in TRAILBACK_PATTERN_DAYS (30) days becomes a Suggested
 *              case (types with a TRAILBACK_PATTERN_SKIP word, like bullying or a weapon, never do).
 *            - Trail Journal flag: a journal submitted with TRAILBACK_TJ_PRIOR (2) or more earlier
 *              journals, or a Low insight journal with at least one before it, becomes Suggested.
 *            - (Teacher referrals and the consequence form's offer box already make cases.)
 *            A student with an open case, or one from the last 30 days, is never suggested again.
 *            The coordinator gets one email listing the new suggestions. An AP still decides to
 *            Offer: that stays a person's call.
 *   2. ASK   - When a case is offered, the family gets a consent link (a one-time page, Yes or No)
 *              if the case has a family email; otherwise the coordinator is asked to call or add it.
 *              Staff can still record consent by hand after a call.
 *   3. JOURNAL - At consent the student gets a Trail Journal link at their school email
 *              (STUDENT_EMAIL_PATTERN, {id}@gomounts.org) and one reminder after 2 school days.
 *   4. SCHEDULE - When the Trail Journal is done, the case goes on the next Wednesday that is at
 *              least 2 school days out, not a day off, with room (PANEL_MAX_CASES, 2). The student,
 *              the family, the referring staff (if an email) and the coordinator are told.
 *   5. SEATS - Scheduling sends a seat call to the Panel Members with the sign-up link; if fewer
 *              than PANEL_MIN_SEATS (3) staff are signed up 2 school days before, a second call.
 * The glimpse: once you sign up, the sign-up page shows the case's strengths, the pattern this
 * year and the Trail Journal status (panelGlimpse_, used by 13_PanelSignup.gs).
 */

// Emails switched off on purpose (TRAILBACK_EMAILS OFF) are not a problem to report.
function quietOff_(list) { return (list || []).filter(function (w) { return !/TRAILBACK_EMAILS is OFF/.test(w); }); }
function pipelineOn_() { return String(webSetting_('TRAILBACK_AUTOMATION') || 'ON').toUpperCase() !== 'OFF'; }
function numSetting_(k, d) { const n = Number(webSetting_(k)); return n > 0 ? n : d; }
function wordsSetting_(k, d) { return String(webSetting_(k) || d).toLowerCase().split(',').map(function (s) { return s.trim(); }).filter(String); }
function firstLastOf_(name) { const m = String(name || '').trim().match(/^([^,]+),\s*(.+)$/); return m ? (m[2] + ' ' + m[1]).replace(/\s+/g, ' ').trim() : String(name || '').trim(); }
function shortName_(name) { const p = String(name || '').trim().split(/\s+/); return p.length > 1 ? p[0] + ' ' + p[p.length - 1].charAt(0) + '.' : (p[0] || 'A student'); }
function studentEmail_(id) { const sid = digits800_(id); return sid ? String(webSetting_('STUDENT_EMAIL_PATTERN') || '{id}@gomounts.org').replace('{id}', sid) : ''; }
function journalUrl_() { return String(webSetting_('SITE_URL') || 'https:\/\/trailjournal.org').replace(/\/+$/, '') + '/'; }

// Students who should not be suggested: an open case, or any case in the last N days.
function recentCaseIds_(days) {
  const since = ymdOf_(new Date(Date.now() - days * 86400000)), out = {};
  allCases_().forEach(function (c) {
    const open = c.status.indexOf('closed') !== 0 && c.status !== 'declined';
    if (c.student_id && (open || String(c.created_at || '').slice(0, 10) >= since || String(c.closed_at || '').slice(0, 10) >= since)) out[c.student_id] = true;
  });
  return out;
}

// ── 1. SPOT ─────────────────────────────────────────────────────────
/** Master Log repeat patterns -> Suggested cases. Returns the new cases. */
function patternSuggestions_() {
  const id = webSetting_('ACCOUNTABILITY_SHEET_ID');
  if (!id) return [];
  const sh = SpreadsheetApp.openById(id).getSheetByName(ML_TAB_);
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getDataRange().getDisplayValues(), h = v[0].map(function (x) { return String(x).trim(); });
  const col = function (n) { return h.indexOf(n); };
  const need = ['Student 800#', 'Behavior Type'];
  if (need.some(function (n) { return col(n) < 0; })) return [];
  const days = numSetting_('TRAILBACK_PATTERN_DAYS', 30), count = numSetting_('TRAILBACK_PATTERN_COUNT', 3);
  const skip = wordsSetting_('TRAILBACK_PATTERN_SKIP', 'bully, threat, weapon, safety, harass, drug, vape, fight');
  const since = ymdOf_(new Date(Date.now() - days * 86400000));
  const by = {};
  v.slice(1).forEach(function (r) {
    const sid = digits800_(r[col('Student 800#')]), type = String(r[col('Behavior Type')] || '').trim();
    const day = ymdOf_(r[col('Consequence Date')] || '') || ymdOf_(r[col('Entry Date')] || '');
    const status = col('Status') > -1 ? String(r[col('Status')]).toLowerCase() : '';
    if (!sid || !type || !day || day < since || /cancel|resched/.test(status)) return;
    if (skip.some(function (w) { return type.toLowerCase().indexOf(w) > -1; })) return;
    const k = sid + '|' + type.toLowerCase();
    (by[k] = by[k] || []).push({ r: r, day: day, type: type });
  });
  const taken = recentCaseIds_(days), made = [];
  Object.keys(by).forEach(function (k) {
    const list = by[k].sort(function (a, b) { return a.day < b.day ? 1 : -1; });
    const sid = k.split('|')[0];
    if (list.length < count || taken[sid]) return;
    const top = list[0].r, det = list.filter(function (x) { return /detention/i.test(String(x.r[col('Consequence Type')] || '')); })[0];
    const c = createCase_({ student_id: sid, student_name: firstLastOf_(top[col('Student Name')]), grade: col('Grade') > -1 ? top[col('Grade')] : '',
      squad: col('Squad') > -1 ? top[col('Squad')] : '', behavior_type: list[0].type, source: 'pattern', status: 'suggested',
      behavior_pattern: list.length + ' ' + list[0].type + ' records in ' + days + ' days (' + list.map(function (x) { return x.day.slice(5); }).join(', ') + ')',
      master_log_row: det ? det.r[col('Entry ID')] : '', consequence_type: det ? det.r[col('Consequence Type')] : '',
      consequence_original_date: det ? det.day : '', incident_summary: col('Reason / Notes') > -1 ? String(top[col('Reason / Notes')] || '').slice(0, 300) : '' }, 'Pattern check');
    if (!c.existing) { taken[sid] = true; made.push(c); }
  });
  return made;
}

/** Trail Journal flags -> Suggested cases. */
function journalSuggestions_() {
  const prior = numSetting_('TRAILBACK_TJ_PRIOR', 2), since = ymdOf_(new Date(Date.now() - 3 * 86400000));
  const taken = recentCaseIds_(30), made = [];
  genericQuery_(TABS.REFLECTIONS, {}).forEach(function (r) {
    const sid = digits800_(r.student_id), when = String(r.created_at || '').slice(0, 10), n = Number(r.prior_count) || 0;
    if (!sid || taken[sid] || String(r.status) !== 'submitted' || when < since) return;
    const low = /low/i.test(String(r.insight_level || ''));
    if (!(n >= prior || (low && n >= 1))) return;
    const c = createCase_({ student_id: sid, student_name: r.student_name || '', grade: r.grade || '', source: 'trail_journal', status: 'suggested',
      behavior_pattern: 'Trail Journal: ' + n + ' earlier journal' + (n === 1 ? '' : 's') + (r.insight_level ? ', insight ' + r.insight_level : '') }, 'Trail Journal flag');
    if (!c.existing) {
      genericUpdate_(TABS.PANEL_CASES, { case_id: c.case_id }, { trail_journal_id: r.session_id || '' });
      taken[sid] = true; made.push(c);
    }
  });
  return made;
}

function suggestionDigest_(cases) {
  if (!cases.length) return null;
  const rows = cases.map(function (c) {
    return emailInfoRow_(escHtml_(c.student_name || c.student_id), escHtml_((c.source === 'pattern' ? 'Repeat pattern: ' : 'Trail Journal: ') + (c.behavior_pattern || '')), squadTheme_(c.squad));
  }).join('');
  return sendTrailbackMail_(coordinatorEmails_(), 'Trailback: ' + cases.length + ' new suggestion' + (cases.length === 1 ? '' : 's'),
    emailShell_('&#128269;', 'New Trailback suggestions', todayYmd_(), rows +
      emailCallout_('Decide', 'Open Trail Journal Admin › Trailback Cases. <strong>Offer</strong> sends the family a consent link; <strong>Decline</strong> clears it.')), false);
}

// ── 2. ASK: the family consent link ─────────────────────────────────
function consentUrl_(token) {
  const base = webSetting_('STUDENT_EXEC_URL') || (function () { try { return ScriptApp.getService().getUrl(); } catch (e) { return ''; } })();
  return base + '?page=consent&t=' + encodeURIComponent(token);
}

/** Sends the consent link (or asks the coordinator to call). Returns warnings. */
function sendConsentLink_(c) {
  try {
    const t = squadTheme_(c.squad);
    if (!emailOf_(c.family_email)) {
      const r = sendTrailbackMail_(coordinatorEmails_(), 'Trailback offered: call the family of ' + (c.student_name || c.case_id),
        emailShell_('&#128222;', 'Offered: family consent needed', escHtml_(c.student_name || c.case_id),
          emailCallout_('No family email on the case', 'Call the family, then record consent on Admin › Trailback Cases. Or add their email on the case (Edit) and use <strong>Send consent link</strong>.', null, t), t), false);
      return r.sent ? ['No family email: the coordinator was asked to call.'] : ['No family email on the case; call the family.'];
    }
    const token = c.consent_token || Utilities.getUuid().replace(/-/g, '');
    genericUpdate_(TABS.PANEL_CASES, { case_id: c.case_id }, { consent_token: token, consent_sent_at: nowIso_() });
    const first = String(c.student_name || '').trim().split(/\s+/)[0] || 'your student';
    const html = emailShell_('&#9978;', 'An invitation for ' + escHtml_(first), 'The Trailback Panel at Ephrata Middle School',
      '<p style="font-family:Verdana,sans-serif;font-size:14px;color:#3a3a3a;line-height:1.6">Instead of serving the detention, ' + escHtml_(first) +
      ' can meet with a small panel of staff (and sometimes trained students) to talk through what happened, hear how it affected others, and make a plan to make it right. ' +
      'If the plan is finished, the detention is cleared. If it isn\'t, the detention is served as planned.</p>' +
      emailCallout_('Your OK', '<a href="' + escHtml_(consentUrl_(token)) + '">Open this page to say yes or no</a>. It takes a minute. Questions? Reply to this email.', null, t), t);
    const r = sendTrailbackMail_(emailOf_(c.family_email), 'Trailback Panel: your OK for ' + first, html, false);
    return r.sent ? [] : ['Consent link not sent: ' + r.why];
  } catch (e) { return ['Consent link not sent: ' + e.message]; }
}

function caseByToken_(token) {
  const t = String(token || '').trim();
  if (t.length < 16) throw new Error('This link is not complete.');
  const r = genericQuery_(TABS.PANEL_CASES, { consent_token: t }, 1)[0];
  if (!r) throw new Error('This link has already been used or has expired. Please call the school.');
  return caseOut_(r);
}

/** The consent page (public, by its one-time link). First name only. */
function consentLookup_(body) {
  const c = caseByToken_((body.payload || {}).t);
  return { first: String(c.student_name || '').trim().split(/\s+/)[0] || 'your student', status: c.status,
    open: c.status === 'offered' || c.status === 'suggested', consequence: String(c.consequence_type || '') };
}

/** { t, answer: 'yes'|'no', name } */
function consentAnswer_(body) {
  const p = body.payload || {}, c = caseByToken_(p.t);
  if (c.status !== 'offered' && c.status !== 'suggested') throw new Error('This has already been answered. Thank you!');
  const who = 'Family via link' + (String(p.name || '').trim() ? ' (' + String(p.name).trim().slice(0, 60) + ')' : '');
  const yes = p.answer === 'yes';
  const r = updateTrailbackCase_({ _staffEmail: who, payload: { case_id: c.case_id, op: yes ? 'consent' : 'decline', reason: yes ? '' : 'Family said no (consent link)', note: who } });
  genericUpdate_(TABS.PANEL_CASES, { case_id: c.case_id }, { consent_token: '' }); // one time only
  return { ok: true, yes: yes, warnings: r.warnings || [] };
}

// ── 3. JOURNAL: the student's Trail Journal link at consent ─────────
function sendJournalLink_(c, reminder) {
  try {
    const to = studentEmail_(c.student_id);
    if (!to) return ['No 800# on the case, so no Trail Journal link was emailed.'];
    const first = String(c.student_name || '').trim().split(/\s+/)[0] || 'there';
    const t = squadTheme_(c.squad);
    const html = emailShell_('&#127956;', reminder ? 'A reminder: your Trail Journal' : 'Your Trail Journal, before the Trailback Panel', 'Hi ' + escHtml_(first),
      '<p style="font-family:Verdana,sans-serif;font-size:14px;color:#3a3a3a;line-height:1.6">Before you meet the Trailback Panel, take about 15 minutes to do a Trail Journal. ' +
      'Use your 800 number so the panel can read it. It helps them hear your side.</p>' +
      emailCallout_('Start here', '<a href="' + escHtml_(journalUrl_()) + '">Open Trail Journal</a>. When you finish, we\'ll let you know your panel date.', null, t), t);
    const r = sendTrailbackMail_(to, reminder ? 'Reminder: your Trail Journal for the Trailback Panel' : 'Your Trail Journal for the Trailback Panel', html, false);
    if (r.sent) genericUpdate_(TABS.PANEL_CASES, { case_id: c.case_id }, reminder ? { journal_reminded_at: nowIso_() } : { journal_sent_at: nowIso_() });
    return r.sent ? [] : ['Trail Journal link not sent: ' + r.why];
  } catch (e) { return ['Trail Journal link not sent: ' + e.message]; }
}

// ── 4. SCHEDULE: the next Wednesday with room ───────────────────────
function nextOpenWednesday_() {
  const max = numSetting_('PANEL_MAX_CASES', 2), off = noSchoolDates_(), today = todayYmd_();
  const taken = {};
  allCases_().forEach(function (c) { if (c.status === 'scheduled' && c.scheduled_date) taken[c.scheduled_date] = (taken[c.scheduled_date] || 0) + 1; });
  let d = nextWednesday_(today);
  for (let i = 0; i < 12; i++) {
    if (!off[d] && schoolDaysBetween_(today, d) >= 2 && (taken[d] || 0) < max) return d;
    const x = new Date(d + 'T12:00:00'); x.setDate(x.getDate() + 7); d = ymdOf_(x);
  }
  return '';
}

function scheduledNotice_(c) {
  const t = squadTheme_(c.squad), who = escHtml_(c.student_name || c.case_id), first = String(c.student_name || '').trim().split(/\s+/)[0] || 'there';
  const d = panelDetails_(c), when = escHtml_(fmtDate_(c.scheduled_date)) + ', ' + escHtml_(d.startTime);
  const out = [];
  const st = studentEmail_(c.student_id);
  if (st) out.push(sendTrailbackMail_(st, 'Your Trailback Panel: ' + fmtDate_(c.scheduled_date),
    emailShell_('&#9978;', 'Your Trailback Panel', 'Hi ' + escHtml_(first), emailInfoRow_('When', when, t) + (d.room ? emailInfoRow_('Where', escHtml_(d.room), t) : '') +
      emailCallout_('What to expect', 'A small group of adults will listen to your side, talk about who was affected, and help you make a plan. Come ready to be honest. You\'ll get a pass.', null, t), t), false));
  if (emailOf_(c.family_email)) out.push(sendTrailbackMail_(emailOf_(c.family_email), 'Trailback Panel date for ' + first,
    emailShell_('&#9978;', 'The Trailback Panel date', escHtml_(first), emailInfoRow_('When', when, t) + emailCallout_('Thank you', 'Thank you for saying yes. We\'ll share the plan after the panel.', null, t), t), false));
  const staff = [emailOf_(c.referring_staff), coordinatorEmails_()].filter(String).join(',');
  if (staff) out.push(sendTrailbackMail_(staff, 'Trailback scheduled: ' + (c.student_name || c.case_id) + ', ' + fmtDate_(c.scheduled_date),
    emailShell_('&#128197;', 'Trailback Panel scheduled', who, emailInfoRow_('When', when, t) + emailCallout_('A pass', 'Please let the student go at the panel time.', null, t), t), false));
  return out;
}

/** Trail Journal done -> on the next open Wednesday. Returns the cases scheduled. */
function autoSchedule_() {
  if (String(webSetting_('TRAILBACK_AUTO_SCHEDULE') || 'ON').toUpperCase() === 'OFF') return [];
  const out = [];
  allCases_().filter(function (c) { return c.status === 'tj_done' && !c.scheduled_date; }).forEach(function (c) {
    const wed = nextOpenWednesday_();
    if (!wed) return;
    try { updateTrailbackCase_({ _staffEmail: 'Auto-schedule', payload: { case_id: c.case_id, op: 'schedule', date: wed, note: 'Scheduled automatically: Trail Journal done' } }); out.push(c.case_id); }
    catch (e) { console.warn('Auto-schedule ' + c.case_id + ': ' + e.message); }
  });
  return out;
}

// ── 5. SEATS ────────────────────────────────────────────────────────
function seatCall_(c, second) {
  try {
    const t = squadTheme_(c.squad), d = panelDetails_(c);
    const on = panelConfirmed_(panelSignupTable_().rows, c.case_id).length;
    const html = emailShell_('&#129517;', second ? 'Still need panelists' : 'Seats open on a Trailback Panel', escHtml_(fmtDate_(c.scheduled_date)) + ', ' + escHtml_(d.startTime),
      emailInfoRow_('Student', escHtml_(shortName_(c.student_name)) + (c.grade ? ' · Grade ' + escHtml_(c.grade) : ''), t) +
      (c.behavior_type ? emailInfoRow_('Behavior', escHtml_(c.behavior_type), t) : '') +
      emailInfoRow_('Seats', on + ' of ' + d.seats + ' taken', t) +
      emailCallout_('Take a seat', '<a href="' + escHtml_(panelSignupUrl_()) + '">Sign up</a>: you get a calendar invite, a glimpse of the student, and the full brief the school day before.', null, t), t);
    const r = sendTrailbackMail_(panelMemberEmails_(), (second ? 'Still need panelists: ' : 'Seats open: ') + 'Trailback Panel ' + fmtDate_(c.scheduled_date), html, false);
    return r.sent ? [] : ['Seat call not sent: ' + r.why];
  } catch (e) { return ['Seat call not sent: ' + e.message]; }
}

function lowSeatCalls_() {
  const min = numSetting_('PANEL_MIN_SEATS', 3), today = todayYmd_(), props = PropertiesService.getScriptProperties(), out = [];
  const signups = panelSignupTable_().rows;
  panelUpcomingCases_().forEach(function (c) {
    if (schoolDaysBetween_(today, c.scheduled_date) !== 2) return;
    if (panelConfirmed_(signups, c.case_id).length >= min || props.getProperty('tb-seats2-' + c.case_id)) return;
    props.setProperty('tb-seats2-' + c.case_id, today);
    out.push.apply(out, seatCall_(c, true));
  });
  return out;
}

// ── THE GLIMPSE (sign-up page, for those signed up) ─────────────────
function panelGlimpse_(c) {
  let b = null;
  try { b = JSON.parse(String(c.climber_brief || '')); } catch (e) { b = null; }
  const st = (b && b.strengths) || {}, pat = (b && b.pattern) || {}, prof = (b && b.profile) || {};
  const j = panelJournalFor_(c);
  return {
    incident: String(c.incident_summary || ''), pattern: String(c.behavior_pattern || ''),
    team: prof.team || '', squad: prof.squad || c.squad || '',
    strengths: (st.counts || []).filter(function (g) { return g.year; }).map(function (g) { return g.label + ': ' + g.year; }).concat(st.activities && st.activities.length ? ['Activities: ' + st.activities.join(', ')] : []),
    yearPattern: (pat.counts || []).map(function (g) { return g.label + ': ' + g.year + ' this year (' + g.recent + ' recently)'; }),
    journal: j ? (j.status === 'submitted' ? 'Trail Journal done ' + j.date : 'Trail Journal started') : 'No Trail Journal yet',
  };
}

// ── TRAIL GUIDES: 60 school days after the Summit, no new Master Log entry ──
function trailGuides_() {
  const days = numSetting_('TRAILBACK_GUIDE_DAYS', 60), today = todayYmd_();
  const due = allCases_().filter(function (c) { return c.status === 'closed_completed' && c.summit_at && !c.guide_eligible_at && c.student_id &&
    schoolDaysBetween_(String(c.summit_at).slice(0, 10), today) >= days; });
  if (!due.length) return [];
  const log = masterLogByStudent_(), out = [];
  due.forEach(function (c) {
    const since = String(c.summit_at).slice(0, 10);
    if ((log[c.student_id] || []).some(function (r) { return r.day > since; })) return; // not yet: check again tomorrow
    genericUpdate_(TABS.PANEL_CASES, { case_id: c.case_id }, { guide_eligible_at: nowIso_() });
    out.push(c);
  });
  if (out.length) sendTrailbackMail_(coordinatorEmails_(), 'Trail Guide: ' + out.length + ' student' + (out.length === 1 ? '' : 's') + ' eligible',
    emailShell_('&#129517;', 'Ready to be a Trail Guide', days + ' school days since the Summit, no new entries',
      out.map(function (c) { return emailInfoRow_(escHtml_(c.student_name || c.student_id), escHtml_((c.grade ? 'Grade ' + c.grade + ' · ' : '') + titleSquad2_(c.squad)), squadTheme_(c.squad)); }).join('') +
      emailCallout_('Invite them', 'They can sit on a future panel as a student panelist: add them to the Panel Members tab with Role "Student".')), false);
  return out.map(function (c) { return c.case_id; });
}

// ── THE DAILY RUN (from trailbackDailyCheck, 12_TrailbackCases.gs) ───
function pipelineDaily_() {
  if (!pipelineOn_()) return { off: true };
  const out = { suggested: [], reminded: 0, scheduled: [], seatCalls: [] };
  try { out.suggested = patternSuggestions_(); } catch (e) { console.warn('Pattern check: ' + e.message); }
  try { out.suggested = out.suggested.concat(journalSuggestions_()); } catch (e) { console.warn('Journal check: ' + e.message); }
  try { suggestionDigest_(out.suggested); } catch (e) { console.warn('Suggestion digest: ' + e.message); }
  // One Trail Journal reminder, 2 school days after consent.
  try {
    allCases_().filter(function (c) { return c.status === 'consented' && c.consent_at && !c.journal_reminded_at && schoolDaysBetween_(String(c.consent_at).slice(0, 10), todayYmd_()) >= 2; })
      .forEach(function (c) { if (!sendJournalLink_(c, true).length) out.reminded++; });
  } catch (e) { console.warn('Journal reminders: ' + e.message); }
  try { out.scheduled = autoSchedule_(); } catch (e) { console.warn('Auto-schedule: ' + e.message); }
  try { out.seatCalls = lowSeatCalls_(); } catch (e) { console.warn('Seat calls: ' + e.message); }
  try { out.guides = trailGuides_(); } catch (e) { console.warn('Trail Guides: ' + e.message); }
  return out;
}

function menuRunPipeline_() {
  const r = pipelineDaily_();
  safeAlert_(r.off ? 'TRAILBACK_AUTOMATION is OFF on the Settings tab.' :
    'Trailback autopilot ran:\n\n' + r.suggested.length + ' new suggestion(s)\n' + r.reminded + ' Trail Journal reminder(s)\n' + r.scheduled.length + ' case(s) scheduled\n' +
    'Seat calls: ' + (r.seatCalls.length ? r.seatCalls.join('; ') : 'none needed'));
}
