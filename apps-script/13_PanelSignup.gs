/**
 * Trail Journal -- Apps Script backend
 * File 13: Trailback Panel sign-up (?page=signup on the staff link), on the live case list.
 *
 * ONE CASE LIST. The sign-up page shows the Trailback cases from panel_cases (12_TrailbackCases.gs)
 * that are Scheduled for a panel date today or later. Cases are never typed twice: a case gets
 * on the page when it is scheduled in the Trailback queue. Panel leads (Settings PANEL_ADMINS)
 * set the panel details on each case (start time, minutes, room, seats, notes, sign-ups open or
 * closed); the student's details come from the case and are not edited here.
 *
 * SIGNING UP. Staff take a seat (panel_signups tab: one row per staff member per case) and get a
 * calendar invite right away (.ics; the district blocks the calendar API, so every panelist gets
 * their own invite for the same time and room). Stepping off sends a cancellation.
 *
 * THE BRIEF, the school day before the panel at about 3:00 (weekends and the No School Dates tab
 * skipped), when TRAILBACK_EMAILS is ON:
 *   - every signed-up panelist gets the case brief: the case, the Climb-er summary (climber_brief,
 *     written onto the case row every hour by Climb-er's Trailback.gs) and the student's newest
 *     Trail Journal;
 *   - cases nobody signed up for go to the Panel Members tab in the summary brief (as before),
 *     with a link to sign up.
 * This replaces the old Tuesday-only brief; the Tuesday timer, if still installed, runs the same.
 *
 * New panel_cases columns (01_Config.gs SCHEMA, added on Run setup): panel_start_time,
 * panel_minutes, panel_room, panel_seats, signups ('open' / 'closed'), panel_notes,
 * climber_brief, climber_brief_at, brief_sent_at.
 */

const PANEL_SIGNUPS_TAB_ = 'panel_signups';
const PANEL_SIGNUP_COLS_ = ['id', 'created_at', 'case_id', 'staff_email', 'staff_name', 'status', 'invite_sent_at', 'withdrawn_at'];
const PANEL_TZ_ = 'America/New_York';
const PANEL_ACTIONS_ = ['panelBoard', 'panelSaveCase', 'panelSignUp', 'panelWithdraw', 'panelSendBrief'];

// ── THE SIGN-UPS TAB ────────────────────────────────────────────────
function panelSignupTable_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PANEL_SIGNUPS_TAB_);
  if (!sheet) {
    sheet = ss.insertSheet(PANEL_SIGNUPS_TAB_);
    sheet.getRange(1, 1, 1, PANEL_SIGNUP_COLS_.length).setValues([PANEL_SIGNUP_COLS_]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  let headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0].map(function (h) { return String(h).trim(); });
  PANEL_SIGNUP_COLS_.forEach(function (c) {
    if (headers.indexOf(c) > -1) return;
    sheet.getRange(1, headers.length + 1).setValue(c);
    headers = headers.concat([c]);
  });
  const values = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues() : [];
  const rows = values.map(function (v, i) { const o = { _row: i + 2 }; headers.forEach(function (h, j) { o[h] = v[j]; }); return o; });
  return {
    rows: rows,
    add: function (o) { sheet.appendRow(headers.map(function (h) { return o[h] == null ? '' : o[h]; })); return o; },
    set: function (row, patch) {
      Object.keys(patch).forEach(function (k) {
        const c = headers.indexOf(k);
        if (c > -1) { sheet.getRange(row._row, c + 1).setValue(patch[k]); row[k] = patch[k]; }
      });
    },
  };
}

function panelHm_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, PANEL_TZ_, 'HH:mm');
  const m = String(v || '').trim().match(/^(\d{1,2}):(\d{2})/);
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : '';
}

// ── WHO IS ASKING ───────────────────────────────────────────────────
function panelStaff_(body) {
  const email = String((body && body._staffEmail) || '').toLowerCase();
  if (!email) throw new Error('Open the sign-up page from the staff link and sign in with your school Google account.');
  return email;
}

function panelAdmins_() {
  const list = String(webSetting_('PANEL_ADMINS') || '').toLowerCase().split(',').map(function (s) { return s.trim(); }).filter(String);
  if (list.length) return list;
  try { return [String(Session.getEffectiveUser().getEmail() || '').toLowerCase()]; } catch (e) { return []; }
}

function panelIsAdmin_(email) { return panelAdmins_().indexOf(String(email).toLowerCase()) > -1; }

function panelRequireAdmin_(body) {
  const email = panelStaff_(body);
  if (!panelIsAdmin_(email)) throw new Error('Only the panel leads can do this. Ask Matt or Kelly.');
  return email;
}

// "kelly_rigg@easdpa.org" -> "Kelly Rigg" (used until the person types their own name).
function panelNameFromEmail_(email) {
  return String(email).split('@')[0].split(/[._]+/).filter(String)
    .map(function (w) { return w.charAt(0).toUpperCase() + w.slice(1); }).join(' ');
}

// ── THE CASES ON THE PAGE ───────────────────────────────────────────
// Scheduled cases from today on.
function panelUpcomingCases_() {
  const today = todayYmd_();
  return allCases_().filter(function (c) { return c.status === 'scheduled' && c.scheduled_date && c.scheduled_date >= today; });
}

function panelFindCase_(caseId) {
  const c = getCase_(caseId);
  if (c.status !== 'scheduled' || !c.scheduled_date) throw new Error('That case is not scheduled for a panel any more.');
  return c;
}

// Panel details with the Settings defaults filled in.
function panelDetails_(c) {
  return {
    startTime: panelHm_(c.panel_start_time) || panelHm_(webSetting_('PANEL_START_TIME')) || '12:00',
    minutes: Number(c.panel_minutes) || Number(webSetting_('PANEL_MINUTES')) || 25,
    room: String(c.panel_room || webSetting_('PANEL_ROOM') || '').trim(),
    seats: Number(c.panel_seats) || Number(webSetting_('PANEL_SEATS')) || 6,
    open: String(c.signups || 'open') !== 'closed',
    notes: String(c.panel_notes || ''),
  };
}

function panelCaseView_(c) {
  const d = panelDetails_(c);
  return {
    id: String(c.case_id), panelDate: c.scheduled_date, startTime: d.startTime, minutes: d.minutes, room: d.room,
    studentName: String(c.student_name || ''), studentId: String(c.student_id || ''), grade: String(c.grade || ''),
    behavior: String(c.behavior_type || c.behavior_pattern || ''), referringStaff: String(c.referring_staff || ''),
    seats: d.seats, status: d.open ? 'open' : 'closed', notes: d.notes, squad: String(c.squad || ''),
    climberReady: !!String(c.climber_brief || '').trim(), briefSentAt: String(c.brief_sent_at || ''),
  };
}

function panelConfirmed_(signups, caseId) {
  return signups.filter(function (s) { return String(s.case_id) === String(caseId) && String(s.status) === 'confirmed'; });
}

// The student's newest Trail Journal: by 800# when the case has one, otherwise by exact name.
function panelJournalFor_(c) {
  const t = recordsTable_();
  const id = String(c.student_id || '').trim(), name = String(c.student_name || '').trim().toLowerCase();
  let best = null;
  for (let i = 0; i < t.values.length; i++) {
    const r = t.row(i);
    const hit = id ? String(r.student_id || '').trim() === id : (name && String(r.student_name || '').trim().toLowerCase() === name);
    if (!hit) continue;
    const when = rowCreatedMs_(r) || 0;
    if (!best || when >= best.when) best = { when: when, row: r };
  }
  if (!best) return null;
  return {
    status: String(best.row.status || ''), text: String(best.row.reflection_text || ''),
    date: best.when ? Utilities.formatDate(new Date(best.when), PANEL_TZ_, 'M/d/yyyy') : '',
  };
}

function panelBoard_(body) {
  const me = panelStaff_(body), admin = panelIsAdmin_(me);
  const signups = panelSignupTable_().rows;
  const out = panelUpcomingCases_().map(panelCaseView_)
    .filter(function (c) { return admin || c.status === 'open'; })
    .sort(function (a, b) { return (a.panelDate + a.startTime).localeCompare(b.panelDate + b.startTime); })
    .map(function (c) {
      const on = panelConfirmed_(signups, c.id);
      c.panelists = on.map(function (s) { return String(s.staff_name || panelNameFromEmail_(s.staff_email)); });
      c.seatsLeft = Math.max(0, c.seats - on.length);
      c.mine = on.some(function (s) { return String(s.staff_email).toLowerCase() === me; });
      const j = panelJournalFor_({ student_id: c.studentId, student_name: c.studentName });
      c.journal = j ? { status: j.status, date: j.date } : null;
      // The glimpse (17_Pipeline.gs): only for those signed up on this case.
      if (c.mine) { try { c.glimpse = panelGlimpse_(getCase_(c.id)); } catch (e) { c.glimpse = null; } }
      return c;
    });
  let queueUrl = '';
  try { queueUrl = String(webSetting_('SITE_URL') || '').replace(/\/+$/, '') + '/panel/'; } catch (e) { queueUrl = ''; }
  return { me: me, myName: panelNameFromEmail_(me), isAdmin: admin, cases: out, queueUrl: queueUrl };
}

// ── WRITE ───────────────────────────────────────────────────────────
/** Panel leads: the panel details on a scheduled case (the student's details are the case's). */
function panelSaveCase_(body) {
  const me = panelRequireAdmin_(body), p = body.payload || {};
  const c = panelFindCase_(p.id);
  const time = panelHm_(p.startTime);
  if (!time) throw new Error('Pick the start time.');
  const patch = {
    panel_start_time: time, panel_minutes: Math.max(5, Math.min(120, Number(p.minutes) || 25)),
    panel_room: String(p.room || '').trim().slice(0, 60), panel_seats: Math.max(1, Math.min(12, Number(p.seats) || 6)),
    signups: p.status === 'closed' ? 'closed' : 'open', panel_notes: String(p.notes || '').trim().slice(0, 600),
  };
  saveCase_(c, patch, me, c.status, 'Panel details: ' + time + ', ' + patch.panel_minutes + ' min' + (patch.panel_room ? ', ' + patch.panel_room : '') +
    ', ' + patch.panel_seats + ' seats, sign-ups ' + patch.signups);
  return { id: String(c.case_id) };
}

function panelSignUp_(body) {
  const me = panelStaff_(body), p = body.payload || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // two people taking the last seat at once
  let c, signup;
  try {
    c = panelFindCase_(p.caseId);
    const d = panelDetails_(c);
    if (!d.open) throw new Error('This case is closed to new sign-ups.');
    if (c.scheduled_date < todayYmd_()) throw new Error('This panel has already met.');
    const t = panelSignupTable_();
    const on = panelConfirmed_(t.rows, c.case_id);
    if (on.some(function (s) { return String(s.staff_email).toLowerCase() === me; })) throw new Error('You are already on this panel.');
    if (on.length >= d.seats) throw new Error('This panel is full.');
    signup = t.add({ id: newUuid_(), created_at: nowIso_(), case_id: String(c.case_id), staff_email: me,
      staff_name: String(p.name || '').trim().slice(0, 60) || panelNameFromEmail_(me), status: 'confirmed', invite_sent_at: '', withdrawn_at: '' });
  } finally { lock.releaseLock(); }

  // The seat is saved even if the email fails; the page shows what happened.
  let invite = 'sent';
  try {
    panelSendInvite_(c, signup);
    const t = panelSignupTable_();
    const row = t.rows.filter(function (r) { return String(r.id) === signup.id; })[0];
    if (row) t.set(row, { invite_sent_at: nowIso_() });
    // Signed up after the brief went out: send it to them now.
    if (String(c.brief_sent_at || '').trim()) panelSendBriefTo_(c, [signup]);
  } catch (err) { invite = 'failed: ' + (err.message || err); }
  return { ok: true, invite: invite };
}

function panelWithdraw_(body) {
  const me = panelStaff_(body), p = body.payload || {};
  const c = getCase_(p.caseId);
  const t = panelSignupTable_();
  const mine = panelConfirmed_(t.rows, c.case_id).filter(function (s) { return String(s.staff_email).toLowerCase() === me; })[0];
  if (!mine) throw new Error('You are not on this panel.');
  t.set(mine, { status: 'withdrawn', withdrawn_at: nowIso_() });
  let note = 'sent';
  try { panelSendCancel_(c, mine); } catch (err) { note = 'failed: ' + (err.message || err); }
  return { ok: true, cancel: note };
}

// ── CALENDAR INVITE ─────────────────────────────────────────────────
function panelTimes_(c) {
  const d = panelDetails_(c), date = c.scheduled_date, hm = d.startTime.split(':').map(Number);
  const end = hm[0] * 60 + hm[1] + d.minutes;
  const eh = Math.floor(end / 60) % 24, em = end % 60;
  return { date: date, startISO: icsDate_(date, hm[0], hm[1]), endISO: icsDate_(date, eh, em),
    label: formatTimeLabel_(hm[0], hm[1]) + ' - ' + formatTimeLabel_(eh, em) };
}

function panelUid_(c, signup) {
  return 'trailback-' + String(c.case_id) + '-' + String(signup.staff_email).replace(/[^a-z0-9]/gi, '') + '@trailjournal.org';
}

function panelIcsText_(s) { return String(s).replace(/\\/g, '\\\\').replace(/\r?\n/g, '\\n').replace(/([,;])/g, '\\$1'); }

function panelSignupUrl_() { return String(webSetting_('SITE_URL') || '').replace(/\/+$/, '') + '/signup/'; }

function panelSendInvite_(c, signup) {
  const organizer = Session.getEffectiveUser().getEmail();
  const tm = panelTimes_(c), d = panelDetails_(c), who = String(c.student_name || c.student_id || 'a student'), room = d.room;
  const sheetUrl = String(webSetting_('PANEL_QUESTION_SHEET_URL') || '').trim();
  const others = panelConfirmed_(panelSignupTable_().rows, c.case_id)
    .filter(function (s) { return String(s.staff_email).toLowerCase() !== String(signup.staff_email).toLowerCase(); })
    .map(function (s) { return String(s.staff_name || panelNameFromEmail_(s.staff_email)); });
  const desc = 'Trailback Panel for ' + who + (c.grade ? ' (grade ' + c.grade + ')' : '') + '.\n' +
    (room ? 'Room: ' + room + '\n' : '') +
    'The case brief (Trail Journal and Climb-er summary) arrives at 3:00 the school day before.\n' +
    'What is said in the room stays in the room.';
  const ics = buildIcsVevent_({ uid: panelUid_(c, signup), startISO: tm.startISO, endISO: tm.endISO,
    summary: panelIcsText_('Trailback Panel: ' + who), description: panelIcsText_(desc),
    organizerEmail: organizer, attendeeEmail: signup.staff_email, attendeeName: signup.staff_name });
  const t = squadTheme_(c.squad);
  const html = emailShell_('🧭', 'You are on the Trailback Panel', escHtml_(who),
    emailInfoRow_('When', escHtml_(fmtDate_(tm.date)) + ', ' + tm.label, t) +
    (room ? emailInfoRow_('Where', escHtml_(room), t) : '') +
    (c.behavior_type ? emailInfoRow_('Behavior', escHtml_(c.behavior_type), t) : '') +
    (others.length ? emailInfoRow_('Also on this panel', escHtml_(others.join(', ')), t) : '') +
    emailCallout_('Before the panel',
      'The case brief arrives at 3:00 the school day before: the student\'s Trail Journal and a Climb-er summary of the pattern and strengths on file.' +
      (sheetUrl ? ' Bring the <a href="' + escHtml_(sheetUrl) + '">panelist question sheet</a>.' : ''), null, t) +
    emailCallout_('If you are close to this student', 'Step off the case on the <a href="' + escHtml_(panelSignupUrl_()) + '">sign-up page</a> and tell Kelly, so an alternate can serve.', { bg: '#F3ECF7' }, t) +
    '<p style="font-family:Verdana,sans-serif;font-size:12px;color:#555">What is said in the room stays in the room. Please do not forward this email.</p>' +
    icsFooterNote_('trailback-panel.ics'), t);
  sendViaMail_({ to: signup.staff_email, subject: 'Trailback Panel: ' + who + ', ' + fmtDate_(tm.date), html: html,
    icsString: ics, icsFilename: 'trailback-panel.ics' });
}

function panelSendCancel_(c, signup) {
  const organizer = Session.getEffectiveUser().getEmail(), tm = panelTimes_(c);
  const who = String(c.student_name || c.student_id || 'a student');
  const ics = 'BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//EMS Trail Journal//EN\r\nCALSCALE:GREGORIAN\r\nMETHOD:CANCEL\r\n' +
    'BEGIN:VEVENT\r\nUID:' + panelUid_(c, signup) + '\r\nSEQUENCE:1\r\nDTSTAMP:' + icsNow_() + '\r\n' +
    'DTSTART;TZID=America/New_York:' + tm.startISO + '\r\nDTEND;TZID=America/New_York:' + tm.endISO + '\r\n' +
    'SUMMARY:' + panelIcsText_('Trailback Panel: ' + who) + '\r\nORGANIZER:mailto:' + organizer + '\r\n' +
    'ATTENDEE:mailto:' + signup.staff_email + '\r\nSTATUS:CANCELLED\r\nEND:VEVENT\r\nEND:VCALENDAR';
  const html = emailShell_('🧭', 'You are off this Trailback Panel', escHtml_(who),
    emailInfoRow_('Was', escHtml_(fmtDate_(tm.date)) + ', ' + tm.label) +
    '<p style="font-family:Verdana,sans-serif;font-size:13px;color:#3a3a3a">Your seat is open again. Open the attached file to take the panel off your calendar.</p>');
  sendViaMail_({ to: signup.staff_email, subject: 'Withdrawn: Trailback Panel, ' + fmtDate_(tm.date), html: html,
    attachments: [Utilities.newBlob(ics, 'text/calendar; charset=UTF-8; method=CANCEL', 'trailback-panel-cancel.ics')] });
}

// ── CASE BRIEF (the case + Climb-er + Trail Journal) ────────────────
function panelList_(items) {
  if (!items || !items.length) return '';
  return '<ul style="margin:6px 0 12px 18px;padding:0;font-family:Verdana,sans-serif;font-size:13px;color:#3a3a3a;line-height:1.6">' +
    items.map(function (x) { return '<li>' + escHtml_(x) + '</li>'; }).join('') + '</ul>';
}

function panelMonth_(ym) {
  const i = Number(String(ym).slice(5, 7)) - 1;
  return ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][i] || String(ym);
}

function panelHead_(text) {
  return '<h3 style="margin:18px 0 6px;font-family:Verdana,sans-serif;font-size:14px;color:' + BRAND_PURPLE_ + ';border-bottom:2px solid ' + BRAND_GOLD_ + ';padding-bottom:4px">' + escHtml_(text) + '</h3>';
}

// climber_brief is JSON written by Climb-er (Trailback.gs there):
// { at, profile: {name, grade, team, squad, homeroom, ceGroup},
//   pattern: {days, total, counts: [{label, year, recent}], months: [{month, count}], records: [{date, source, headline}]},
//   strengths: {counts: [{label, year}], recent: [...], activities: [..]} }
function panelClimberHtml_(c) {
  let b = null;
  try { b = JSON.parse(String(c.climber_brief || '')); } catch (e) { b = null; }
  if (!b || b.error) {
    return panelHead_('From Climb-er') + '<p style="font-family:Verdana,sans-serif;font-size:13px;color:#3a3a3a">' +
      escHtml_(b && b.error ? 'Climb-er could not build a summary: ' + b.error : 'The Climb-er summary is not on this case yet. It fills in every hour once the case has the student\'s 800#.') + '</p>';
  }
  const p = b.profile || {}, pat = b.pattern || {}, st = b.strengths || {};
  const line = function (r) { return [r.date, r.source, r.headline].filter(String).join(' · '); };
  const who = [p.squad ? p.squad + ' squad' : '', p.team ? 'Team ' + p.team : '', p.homeroom ? 'Homeroom ' + p.homeroom : '',
    p.ceGroup ? 'Core Extension: ' + p.ceGroup : ''].filter(String);
  return panelHead_('Squad, team and schedule') + panelList_(who.length ? who : ['Nothing on file.']) +
    panelHead_('Strengths on file') +
    panelList_((st.counts || []).filter(function (g) { return g.year; }).map(function (g) { return g.label + ': ' + g.year + ' this year'; })
      .concat((st.activities || []).length ? ['Activities: ' + st.activities.join(', ')] : [])
      .concat((st.recent || []).map(line))) +
    (!(st.counts || []).some(function (g) { return g.year; }) && !(st.recent || []).length && !(st.activities || []).length
      ? '<p style="font-family:Verdana,sans-serif;font-size:13px;color:#3a3a3a">No shout-outs or activities on file yet. The Strengths stage matters more for this student.</p>' : '') +
    panelHead_('Pattern this school year') +
    panelList_((pat.counts || []).map(function (g) { return g.label + ': ' + g.year + ' this year (' + g.recent + ' in the last ' + (pat.days || 42) + ' days)'; })) +
    ((pat.months || []).length ? '<p style="font-family:Verdana,sans-serif;font-size:13px;color:#3a3a3a;margin:0 0 8px"><strong>By month:</strong> ' +
      escHtml_(pat.months.map(function (m) { return panelMonth_(m.month) + ' ' + m.count; }).join(' · ')) + '</p>' : '') +
    panelList_((pat.records || []).map(line)) +
    ((pat.total || 0) > (pat.records || []).length ? '<p style="font-family:Verdana,sans-serif;font-size:12px;color:#777">Showing the newest ' + pat.records.length + ' of ' + pat.total + ' records. The rest are on the student\'s Climb-er page.</p>' : '') +
    '<p style="font-family:Verdana,sans-serif;font-size:11px;color:#777">Climb-er summary as of ' + escHtml_(String(b.at || '').slice(0, 16).replace('T', ' ')) + '.</p>';
}

function panelBriefHtml_(c) {
  const tm = panelTimes_(c), d = panelDetails_(c), j = panelJournalFor_(c), t = squadTheme_(c.squad);
  const sheetUrl = String(webSetting_('PANEL_QUESTION_SHEET_URL') || '').trim();
  const deadline = webSetting_('TRAILBACK_CONFLICT_DEADLINE') || '3:00 today';
  const journal = j && String(j.text).trim()
    ? '<p style="font-family:Verdana,sans-serif;font-size:11px;color:#777">Written ' + escHtml_(j.date) + (j.status !== 'submitted' ? ' (not finished)' : '') + '.</p>' + formatSummaryHtml_(j.text, t)
    : '<p style="font-family:Verdana,sans-serif;font-size:13px;color:#3a3a3a">' + (j ? 'The student has started a Trail Journal but the report is not ready yet.' : 'No Trail Journal found for this student yet.') + '</p>';
  return emailShell_('📋', 'Trailback Panel case brief', escHtml_(String(c.student_name || c.student_id || '')),
    emailInfoRow_('When', escHtml_(fmtDate_(tm.date)) + ', ' + tm.label, t) +
    (d.room ? emailInfoRow_('Where', escHtml_(d.room), t) : '') +
    (c.grade ? emailInfoRow_('Grade', escHtml_(String(c.grade)), t) : '') +
    (c.behavior_type ? emailInfoRow_('Behavior', escHtml_(String(c.behavior_type)), t) : '') +
    (c.incident_summary ? emailInfoRow_('Incident', escHtml_(String(c.incident_summary)), t) : '') +
    (c.behavior_pattern ? emailInfoRow_('Pattern', escHtml_(String(c.behavior_pattern)), t) : '') +
    (c.referring_staff ? emailInfoRow_('Referred by', escHtml_(String(c.referring_staff)), t) : '') +
    (d.notes ? emailInfoRow_('Notes', escHtml_(d.notes), t) : '') +
    emailCallout_('Read before the panel', 'About five minutes. Start with the strengths, then the student\'s own words.' +
      (sheetUrl ? ' Questions are on the <a href="' + escHtml_(sheetUrl) + '">panelist question sheet</a>.' : '') +
      ' If you are close to this student or the people involved, tell Kelly by <strong>' + escHtml_(deadline) + '</strong> so an alternate can serve.', null, t) +
    panelClimberHtml_(c) +
    panelHead_('The student\'s Trail Journal') + journal +
    '<p style="font-family:Verdana,sans-serif;font-size:12px;color:#555;margin-top:16px">Confidential. Please do not forward or print this email.</p>', t);
}

function panelSendBriefTo_(c, signups) {
  if (!signups.length) return 0;
  const html = panelBriefHtml_(c), tm = panelTimes_(c);
  signups.forEach(function (s) {
    sendViaMail_({ to: s.staff_email, subject: 'Case brief: Trailback Panel, ' + (c.student_name || c.student_id) + ', ' + fmtDate_(tm.date), html: html });
  });
  return signups.length;
}

function panelSendBriefForCase_(caseId) {
  const c = getCase_(caseId);
  const on = panelConfirmed_(panelSignupTable_().rows, c.case_id);
  const sent = panelSendBriefTo_(c, on);
  if (sent) genericUpdate_(TABS.PANEL_CASES, { case_id: c.case_id }, { brief_sent_at: nowIso_() });
  return { sent: sent };
}

function panelSendBrief_(body) {
  panelRequireAdmin_(body);
  return panelSendBriefForCase_((body.payload || {}).caseId);
}

// ── THE SCHOOL DAY BEFORE, ABOUT 3:00 ───────────────────────────────
// The last school day before ymd (weekends and the No School Dates tab skipped).
function panelSchoolDayBefore_(ymd) {
  const off = noSchoolDates_();
  const d = new Date(ymd + 'T12:00:00');
  for (let i = 0; i < 30; i++) {
    d.setDate(d.getDate() - 1);
    const s = ymdOf_(d);
    if (d.getDay() !== 0 && d.getDay() !== 6 && !off[s]) return s;
  }
  return ymdOf_(d);
}

/**
 * Every brief that is due: a scheduled case whose brief day (the school day before the panel) is
 * today, or has passed without one going out. Signed-up panelists get the case brief; cases with
 * nobody signed up go to the Panel Members in the summary brief, once per panel date.
 * preview = build and return it, send nothing.
 */
function panelDailyBriefs_(preview) {
  const today = todayYmd_(), out = { briefs: [], fallback: null };
  if (!preview && !trailbackEmailsOn_()) return { skipped: 'TRAILBACK_EMAILS is OFF on the Settings tab.' };
  const signups = panelSignupTable_().rows, empty = {};
  panelUpcomingCases_().forEach(function (c) {
    if (c.scheduled_date <= today || panelSchoolDayBefore_(c.scheduled_date) > today || String(c.brief_sent_at || '').trim()) return;
    const on = panelConfirmed_(signups, c.case_id);
    if (!on.length) { (empty[c.scheduled_date] = empty[c.scheduled_date] || []).push(c); return; }
    if (preview) { out.briefs.push({ case_id: c.case_id, to: on.map(function (s) { return s.staff_email; }).join(','), html: panelBriefHtml_(c) }); return; }
    try { panelSendBriefForCase_(c.case_id); out.briefs.push({ case_id: c.case_id, sent: on.length }); }
    catch (err) { console.warn('Trailback brief failed for case ' + c.case_id + ': ' + (err.message || err)); }
  });
  const props = PropertiesService.getScriptProperties();
  Object.keys(empty).forEach(function (date) {
    if (!preview && props.getProperty('panel-fallback-' + date)) return;
    const note = emailCallout_('Nobody has signed up yet', 'Take a seat on the <a href="' + escHtml_(panelSignupUrl_()) + '">sign-up page</a>: you get a calendar invite and the full case brief.');
    const r = sendTrailbackMail_(panelMemberEmails_(), 'Trailback Panel brief: ' + date + ' (sign-ups needed)', briefHtml_(date, empty[date], note), preview);
    if (!preview && r.sent) props.setProperty('panel-fallback-' + date, nowIso_());
    out.fallback = r;
  });
  return out;
}

function panelDailyBriefs() {
  const dow = new Date().getDay();
  if (dow === 0 || dow === 6) return;
  panelDailyBriefs_(false);
}

/** Menu: the daily 3:00 brief timer (replaces the Tuesday-only one). */
function setupPanelBriefTrigger() {
  ['panelDailyBriefs', 'trailbackTuesdayBrief'].forEach(function (fn) {
    ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === fn) ScriptApp.deleteTrigger(t); });
  });
  ScriptApp.newTrigger('panelDailyBriefs').timeBased().everyDays(1).atHour(15).inTimezone(PANEL_TZ_).create();
  ensureSchemaTab_(TABS.PANEL_CASES);
  panelSignupTable_();
  safeAlert_('Trailback case briefs are on: every school day at about 3:00 PM, panelists signed up for the next school day\'s panel get the case brief ' +
    '(Trail Journal + Climb-er summary). Cases nobody signed up for go to the Panel Members, as before. The Tuesday-only timer was replaced.\n\n' +
    'Emails only go out when Settings TRAILBACK_EMAILS is ON (it is ' + (trailbackEmailsOn_() ? 'ON' : 'OFF') + ').');
}

// ── MENU ────────────────────────────────────────────────────────────
function previewPanelBriefs() {
  const r = panelDailyBriefs_(true);
  const parts = (r.briefs || []).map(function (b) { return '<p style="font-family:Verdana;font-size:12px"><b>To:</b> ' + escHtml_(b.to) + '</p>' + b.html; });
  if (r.fallback) parts.push('<p style="font-family:Verdana;font-size:12px"><b>To the Panel Members (nobody signed up):</b> ' + escHtml_(r.fallback.to || '(nobody on the tab)') + '</p>' + (r.fallback.html || ''));
  const html = parts.length ? parts.join('<hr>') : '<p style="font-family:Verdana">No brief is due today: no scheduled case is on the next school day, or its brief already went out.</p>';
  try { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(680).setHeight(640), 'Case briefs due today (nothing sent)'); }
  catch (e) { Logger.log(html); }
}

function showPanelSignupLink() {
  let staff = '';
  try { staff = ScriptApp.getService().getUrl() + '?page=signup'; } catch (e) { staff = ''; }
  safeAlert_('The Trailback Panel sign-up page:\n\n' + panelSignupUrl_() + '\n\n(or the staff link with ?page=signup: ' + staff + ')\n\n' +
    'Cases show up there once they are Scheduled in the Trailback queue. Panel leads (Settings PANEL_ADMINS) set the time, room and seats on each case.');
}
