/**
 * Trail Journal -- Apps Script backend
 * File 11: Trailback Panel sign-up (?page=signup on the staff link).
 *
 * Staff open the sign-up page, see the upcoming panel cases and take a seat on one. Signing up
 * sends them a calendar invite straight away. At 3:00 on the school day before the panel every
 * panelist on the case gets the case brief: the student's Trail Journal and the Climb-er summary.
 *
 * Two tabs, made the first time they are needed:
 *   panel_cases    one row per case (student, date, time, room, seats). Climb-er fills
 *                  climber_brief on these rows (Trailback.gs in the Climb-er project); this
 *                  project never opens Climb-er.
 *   panel_signups  one row per staff member per case (confirmed / withdrawn).
 *
 * Only staff signed in on the staff link can use any of this (their school account is the
 * identity on the sign-up). Adding cases and sending briefs is limited to Settings >
 * PANEL_ADMINS (comma separated emails); blank = the account that owns this script.
 *
 * The district blocks the Microsoft calendar API, so there is no single shared calendar event:
 * each panelist gets their own invite (.ics) for the same time and room, like the follow-up
 * check-ins in 04_FollowupEmails.gs.
 */

const PANEL_CASES_TAB_ = 'panel_cases';
const PANEL_SIGNUPS_TAB_ = 'panel_signups';
const PANEL_CASE_COLS_ = ['id', 'created_at', 'created_by', 'panel_date', 'start_time', 'minutes', 'room',
  'student_name', 'student_id', 'grade', 'behavior', 'referring_staff', 'seats', 'status', 'notes',
  'climber_brief', 'climber_brief_at', 'brief_sent_at'];
const PANEL_SIGNUP_COLS_ = ['id', 'created_at', 'case_id', 'staff_email', 'staff_name', 'status', 'invite_sent_at', 'withdrawn_at'];
const PANEL_TZ_ = 'America/New_York';
const PANEL_ACTIONS_ = ['panelBoard', 'panelSaveCase', 'panelSignUp', 'panelWithdraw', 'panelSendBrief'];

// ── SHEET ACCESS ────────────────────────────────────────────────────
// A tab as a list of row objects. Dates and times are kept as text (yyyy-MM-dd, HH:mm).
function panelTable_(tab, cols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(tab);
  if (!sheet) {
    sheet = ss.insertSheet(tab);
    sheet.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sheet.getRange(1, 1, sheet.getMaxRows(), cols.length).setNumberFormat('@'); // keep dates and times as typed
    sheet.setFrozenRows(1);
  }
  let headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0].map(function (h) { return String(h).trim(); });
  cols.forEach(function (c) {
    if (headers.indexOf(c) > -1) return;
    sheet.getRange(1, headers.length + 1).setValue(c);
    headers = headers.concat([c]);
  });
  const values = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues() : [];
  const rows = values.map(function (v, i) {
    const o = { _row: i + 2 };
    headers.forEach(function (h, j) { o[h] = v[j]; });
    return o;
  });
  return {
    rows: rows,
    add: function (o) {
      sheet.appendRow(headers.map(function (h) { return o[h] == null ? '' : o[h]; }));
      return o;
    },
    set: function (row, patch) {
      Object.keys(patch).forEach(function (k) {
        const c = headers.indexOf(k);
        if (c > -1) { sheet.getRange(row._row, c + 1).setValue(patch[k]); row[k] = patch[k]; }
      });
    },
  };
}

function panelYmd_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, PANEL_TZ_, 'yyyy-MM-dd');
  const m = String(v || '').trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : '';
}

function panelHm_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, PANEL_TZ_, 'HH:mm');
  const m = String(v || '').trim().match(/^(\d{1,2}):(\d{2})/);
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : '';
}

function panelToday_() { return Utilities.formatDate(new Date(), PANEL_TZ_, 'yyyy-MM-dd'); }

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

// ── READ ────────────────────────────────────────────────────────────
function panelCaseView_(c) {
  return {
    id: String(c.id), panelDate: panelYmd_(c.panel_date), startTime: panelHm_(c.start_time), minutes: Number(c.minutes) || 20,
    room: String(c.room || ''), studentName: String(c.student_name || ''), studentId: String(c.student_id || ''),
    grade: String(c.grade || ''), behavior: String(c.behavior || ''), referringStaff: String(c.referring_staff || ''),
    seats: Number(c.seats) || 6, status: String(c.status || 'open'), notes: String(c.notes || ''),
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
  const cases = panelTable_(PANEL_CASES_TAB_, PANEL_CASE_COLS_).rows;
  const signups = panelTable_(PANEL_SIGNUPS_TAB_, PANEL_SIGNUP_COLS_).rows;
  const today = panelToday_();
  const out = cases.map(panelCaseView_)
    .filter(function (c) { return c.id && c.panelDate >= today && (admin || c.status === 'open'); })
    .sort(function (a, b) { return (a.panelDate + a.startTime).localeCompare(b.panelDate + b.startTime); })
    .map(function (c) {
      const on = panelConfirmed_(signups, c.id);
      c.panelists = on.map(function (s) { return String(s.staff_name || panelNameFromEmail_(s.staff_email)); });
      c.seatsLeft = Math.max(0, c.seats - on.length);
      c.mine = on.some(function (s) { return String(s.staff_email).toLowerCase() === me; });
      const j = panelJournalFor_({ student_id: c.studentId, student_name: c.studentName });
      c.journal = j ? { status: j.status, date: j.date } : null;
      return c;
    });
  return { me: me, myName: panelNameFromEmail_(me), isAdmin: admin, cases: out };
}

// ── WRITE ───────────────────────────────────────────────────────────
function panelSaveCase_(body) {
  const me = panelRequireAdmin_(body), p = body.payload || {};
  const date = panelYmd_(p.panelDate), time = panelHm_(p.startTime);
  if (!String(p.studentName || '').trim()) throw new Error('Type the student\'s name.');
  if (!date || !time) throw new Error('Pick the panel date and start time.');
  const t = panelTable_(PANEL_CASES_TAB_, PANEL_CASE_COLS_);
  const fields = {
    panel_date: date, start_time: time, minutes: Number(p.minutes) || 20, room: String(p.room || '').trim(),
    student_name: String(p.studentName).trim(), student_id: String(p.studentId || '').replace(/\D/g, ''),
    grade: String(p.grade || '').trim(), behavior: String(p.behavior || '').trim(),
    referring_staff: String(p.referringStaff || '').trim(), seats: Math.max(1, Number(p.seats) || 6),
    status: p.status === 'closed' ? 'closed' : 'open', notes: String(p.notes || '').trim(),
  };
  if (p.id) {
    const row = t.rows.filter(function (r) { return String(r.id) === String(p.id); })[0];
    if (!row) throw new Error('That case is no longer on the list.');
    // A new student on the row means the old Climb-er summary no longer applies.
    if (String(row.student_id) !== fields.student_id || String(row.student_name) !== fields.student_name) {
      fields.climber_brief = ''; fields.climber_brief_at = '';
    }
    t.set(row, fields);
    return { id: String(row.id) };
  }
  const id = newUuid_();
  t.add(Object.assign({ id: id, created_at: nowIso_(), created_by: me }, fields));
  return { id: id };
}

function panelFindCase_(caseId) {
  const t = panelTable_(PANEL_CASES_TAB_, PANEL_CASE_COLS_);
  const row = t.rows.filter(function (r) { return String(r.id) === String(caseId); })[0];
  if (!row) throw new Error('That case is no longer on the list.');
  return { table: t, row: row };
}

function panelSignUp_(body) {
  const me = panelStaff_(body), p = body.payload || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // two people taking the last seat at once
  let c, signup;
  try {
    c = panelFindCase_(p.caseId).row;
    if (String(c.status || 'open') !== 'open') throw new Error('This case is closed to new sign-ups.');
    if (panelYmd_(c.panel_date) < panelToday_()) throw new Error('This panel has already met.');
    const t = panelTable_(PANEL_SIGNUPS_TAB_, PANEL_SIGNUP_COLS_);
    const on = panelConfirmed_(t.rows, c.id);
    if (on.some(function (s) { return String(s.staff_email).toLowerCase() === me; })) throw new Error('You are already on this panel.');
    if (on.length >= (Number(c.seats) || 6)) throw new Error('This panel is full.');
    signup = t.add({ id: newUuid_(), created_at: nowIso_(), case_id: String(c.id), staff_email: me,
      staff_name: String(p.name || '').trim() || panelNameFromEmail_(me), status: 'confirmed', invite_sent_at: '', withdrawn_at: '' });
  } finally { lock.releaseLock(); }

  // The seat is saved even if the email fails; the page shows what happened.
  let invite = 'sent';
  try {
    panelSendInvite_(c, signup);
    const t = panelTable_(PANEL_SIGNUPS_TAB_, PANEL_SIGNUP_COLS_);
    const row = t.rows.filter(function (r) { return String(r.id) === signup.id; })[0];
    if (row) t.set(row, { invite_sent_at: nowIso_() });
    // Signed up after the brief went out: send it to them now.
    if (String(c.brief_sent_at || '').trim()) panelSendBriefTo_(c, [signup]);
  } catch (err) { invite = 'failed: ' + (err.message || err); }
  return { ok: true, invite: invite };
}

function panelWithdraw_(body) {
  const me = panelStaff_(body), p = body.payload || {};
  const c = panelFindCase_(p.caseId).row;
  const t = panelTable_(PANEL_SIGNUPS_TAB_, PANEL_SIGNUP_COLS_);
  const mine = panelConfirmed_(t.rows, c.id).filter(function (s) { return String(s.staff_email).toLowerCase() === me; })[0];
  if (!mine) throw new Error('You are not on this panel.');
  t.set(mine, { status: 'withdrawn', withdrawn_at: nowIso_() });
  let note = 'sent';
  try { panelSendCancel_(c, mine); } catch (err) { note = 'failed: ' + (err.message || err); }
  return { ok: true, cancel: note };
}

// ── CALENDAR INVITE ─────────────────────────────────────────────────
function panelTimes_(c) {
  const date = panelYmd_(c.panel_date), hm = panelHm_(c.start_time).split(':').map(Number);
  const mins = Number(c.minutes) || 20, end = hm[0] * 60 + hm[1] + mins;
  const eh = Math.floor(end / 60) % 24, em = end % 60;
  return { date: date, startISO: icsDate_(date, hm[0], hm[1]), endISO: icsDate_(date, eh, em),
    label: formatTimeLabel_(hm[0], hm[1]) + ' - ' + formatTimeLabel_(eh, em) };
}

function panelUid_(c, signup) {
  return 'trailback-' + String(c.id) + '-' + String(signup.staff_email).replace(/[^a-z0-9]/gi, '') + '@trailjournal.org';
}

function panelIcsText_(s) { return String(s).replace(/\\/g, '\\\\').replace(/\r?\n/g, '\\n').replace(/([,;])/g, '\\$1'); }

function panelSendInvite_(c, signup) {
  const organizer = Session.getEffectiveUser().getEmail();
  const tm = panelTimes_(c), who = String(c.student_name), room = String(c.room || '').trim();
  const sheetUrl = String(webSetting_('PANEL_QUESTION_SHEET_URL') || '').trim();
  const signupUrl = webSetting_('SITE_URL').replace(/\/+$/, '') + '/signup/';
  const others = panelConfirmed_(panelTable_(PANEL_SIGNUPS_TAB_, PANEL_SIGNUP_COLS_).rows, c.id)
    .filter(function (s) { return String(s.staff_email).toLowerCase() !== String(signup.staff_email).toLowerCase(); })
    .map(function (s) { return String(s.staff_name || panelNameFromEmail_(s.staff_email)); });
  const desc = 'Trailback Panel for ' + who + (c.grade ? ' (grade ' + c.grade + ')' : '') + '.\n' +
    (room ? 'Room: ' + room + '\n' : '') +
    'The case brief (Trail Journal and Climb-er summary) arrives at 3:00 the school day before.\n' +
    'What is said in the room stays in the room.';
  const ics = buildIcsVevent_({ uid: panelUid_(c, signup), startISO: tm.startISO, endISO: tm.endISO,
    summary: panelIcsText_('Trailback Panel: ' + who), description: panelIcsText_(desc),
    organizerEmail: organizer, attendeeEmail: signup.staff_email, attendeeName: signup.staff_name });
  const html = emailShell_('🧭', 'You are on the Trailback Panel', escHtml_(who),
    emailInfoRow_('When', escHtml_(fmtDate_(tm.date)) + ', ' + tm.label) +
    (room ? emailInfoRow_('Where', escHtml_(room)) : '') +
    (c.behavior ? emailInfoRow_('Behavior', escHtml_(c.behavior)) : '') +
    (others.length ? emailInfoRow_('Also on this panel', escHtml_(others.join(', '))) : '') +
    emailCallout_('Before the panel',
      'The case brief arrives at 3:00 the school day before: the student\'s Trail Journal and a Climb-er summary of the pattern and strengths on file.' +
      (sheetUrl ? ' Bring the <a href="' + escHtml_(sheetUrl) + '">panelist question sheet</a>.' : '')) +
    emailCallout_('If you are close to this student', 'Step off the case on the <a href="' + escHtml_(signupUrl) + '">sign-up page</a> and tell Kelly, so an alternate can serve.', { bg: '#F3ECF7' }) +
    '<p style="font-family:Verdana,sans-serif;font-size:12px;color:#555">What is said in the room stays in the room. Please do not forward this email.</p>' +
    icsFooterNote_('trailback-panel.ics'));
  sendViaMail_({ to: signup.staff_email, subject: 'Trailback Panel: ' + who + ', ' + fmtDate_(tm.date), html: html,
    icsString: ics, icsFilename: 'trailback-panel.ics' });
}

function panelSendCancel_(c, signup) {
  const organizer = Session.getEffectiveUser().getEmail(), tm = panelTimes_(c);
  const ics = 'BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//EMS Trail Journal//EN\r\nCALSCALE:GREGORIAN\r\nMETHOD:CANCEL\r\n' +
    'BEGIN:VEVENT\r\nUID:' + panelUid_(c, signup) + '\r\nSEQUENCE:1\r\nDTSTAMP:' + icsNow_() + '\r\n' +
    'DTSTART;TZID=America/New_York:' + tm.startISO + '\r\nDTEND;TZID=America/New_York:' + tm.endISO + '\r\n' +
    'SUMMARY:' + panelIcsText_('Trailback Panel: ' + c.student_name) + '\r\nORGANIZER:mailto:' + organizer + '\r\n' +
    'ATTENDEE:mailto:' + signup.staff_email + '\r\nSTATUS:CANCELLED\r\nEND:VEVENT\r\nEND:VCALENDAR';
  const html = emailShell_('🧭', 'You are off this Trailback Panel', escHtml_(String(c.student_name)),
    emailInfoRow_('Was', escHtml_(fmtDate_(tm.date)) + ', ' + tm.label) +
    '<p style="font-family:Verdana,sans-serif;font-size:13px;color:#3a3a3a">Your seat is open again. Open the attached file to take the panel off your calendar.</p>');
  sendViaMail_({ to: signup.staff_email, subject: 'Withdrawn: Trailback Panel, ' + fmtDate_(tm.date), html: html,
    attachments: [Utilities.newBlob(ics, 'text/calendar; charset=UTF-8; method=CANCEL', 'trailback-panel-cancel.ics')] });
}

// ── CASE BRIEF (Trail Journal + Climb-er) ───────────────────────────
function panelList_(items) {
  if (!items || !items.length) return '';
  return '<ul style="margin:6px 0 12px 18px;padding:0;font-family:Verdana,sans-serif;font-size:13px;color:#3a3a3a;line-height:1.6">' +
    items.map(function (x) { return '<li>' + escHtml_(x) + '</li>'; }).join('') + '</ul>';
}

// "2026-09" -> "Sep"
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
// The pattern is the whole school year; `recent` is how many of the year's records fall in the last `days` days.
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
  const tm = panelTimes_(c), j = panelJournalFor_(c);
  const sheetUrl = String(webSetting_('PANEL_QUESTION_SHEET_URL') || '').trim();
  const journal = j && String(j.text).trim()
    ? '<p style="font-family:Verdana,sans-serif;font-size:11px;color:#777">Written ' + escHtml_(j.date) + (j.status !== 'submitted' ? ' (not finished)' : '') + '.</p>' + formatSummaryHtml_(j.text)
    : '<p style="font-family:Verdana,sans-serif;font-size:13px;color:#3a3a3a">' + (j ? 'The student has started a Trail Journal but the report is not ready yet.' : 'No Trail Journal found for this student yet.') + '</p>';
  return emailShell_('📋', 'Trailback Panel case brief', escHtml_(String(c.student_name)),
    emailInfoRow_('When', escHtml_(fmtDate_(tm.date)) + ', ' + tm.label) +
    (c.room ? emailInfoRow_('Where', escHtml_(String(c.room))) : '') +
    (c.grade ? emailInfoRow_('Grade', escHtml_(String(c.grade))) : '') +
    (c.behavior ? emailInfoRow_('Behavior', escHtml_(String(c.behavior))) : '') +
    (c.referring_staff ? emailInfoRow_('Referred by', escHtml_(String(c.referring_staff))) : '') +
    (c.notes ? emailInfoRow_('Notes', escHtml_(String(c.notes))) : '') +
    emailCallout_('Read before the panel', 'About five minutes. Start with the strengths, then the student\'s own words.' +
      (sheetUrl ? ' Questions are on the <a href="' + escHtml_(sheetUrl) + '">panelist question sheet</a>.' : '') +
      ' If you are close to this student, tell Kelly by 3:00 so an alternate can serve.') +
    panelClimberHtml_(c) +
    panelHead_('The student\'s Trail Journal') + journal +
    '<p style="font-family:Verdana,sans-serif;font-size:12px;color:#555;margin-top:16px">Confidential. Please do not forward or print this email.</p>');
}

function panelSendBriefTo_(c, signups) {
  if (!signups.length) return 0;
  const html = panelBriefHtml_(c), tm = panelTimes_(c);
  signups.forEach(function (s) {
    sendViaMail_({ to: s.staff_email, subject: 'Case brief: Trailback Panel, ' + c.student_name + ', ' + fmtDate_(tm.date), html: html });
  });
  return signups.length;
}

function panelSendBriefForCase_(caseId) {
  const found = panelFindCase_(caseId), c = found.row;
  const on = panelConfirmed_(panelTable_(PANEL_SIGNUPS_TAB_, PANEL_SIGNUP_COLS_).rows, c.id);
  const sent = panelSendBriefTo_(c, on);
  if (sent) found.table.set(c, { brief_sent_at: nowIso_() });
  return { sent: sent };
}

function panelSendBrief_(body) {
  panelRequireAdmin_(body);
  return panelSendBriefForCase_((body.payload || {}).caseId);
}

// ── 3:00 THE SCHOOL DAY BEFORE ──────────────────────────────────────
// Tomorrow, or Monday when today is Friday. Days off are not known here: a panel on the first
// day back gets its brief on the last weekday before, which may be a day off.
function panelNextWeekday_(ymd) {
  const d = new Date(ymd + 'T12:00:00');
  do { d.setDate(d.getDate() + 1); } while (d.getDay() === 0 || d.getDay() === 6);
  return Utilities.formatDate(d, PANEL_TZ_, 'yyyy-MM-dd');
}

function panelDailyBriefs() {
  const target = panelNextWeekday_(panelToday_());
  const cases = panelTable_(PANEL_CASES_TAB_, PANEL_CASE_COLS_).rows;
  let sent = 0;
  cases.forEach(function (c) {
    if (panelYmd_(c.panel_date) !== target || String(c.status || 'open') === 'closed' || String(c.brief_sent_at || '').trim()) return;
    try { sent += panelSendBriefForCase_(c.id).sent; } catch (err) { console.warn('Trailback brief failed for case ' + c.id + ': ' + (err.message || err)); }
  });
  return sent;
}

function setupPanelBriefTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'panelDailyBriefs') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('panelDailyBriefs').timeBased().everyDays(1).atHour(15).inTimezone(PANEL_TZ_).create();
  panelTable_(PANEL_CASES_TAB_, PANEL_CASE_COLS_);
  panelTable_(PANEL_SIGNUPS_TAB_, PANEL_SIGNUP_COLS_);
  safeAlert_('Trailback case briefs are on: every day at about 3:00 PM, panelists on the next school day\'s cases get the Trail Journal and Climb-er summary. The panel_cases and panel_signups tabs are ready.');
}
