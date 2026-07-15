/**
 * Trail Journal -- Apps Script backend
 * File 4 of N: Follow-up check-in emails, the one-time plan-summary email, and
 * Outlook-compatible calendar invites (.ics).
 *
 * sendFollowupEmails_ is a faithful port of the live Supabase edge function
 * `send-followup-emails` (v1) -- same three-interval schedule (2-day / 2-week / 2-month),
 * same grade-based time windows, same hand-built .ics attachment so staff on Outlook get a
 * normal calendar invite.
 *
 * sendSingleFollowupEvent_ sends exactly one custom-dated invite (from the main journal's
 * "Schedule Follow-Up Check-In" form) with the full AI plan summary bundled into the same
 * email, replacing what used to be two separate emails.
 *
 * TRANSPORT: this used to go through Resend (a third-party email API), which needed its own
 * API key and a verified sending domain -- and even when Resend accepted the send, the mail
 * still landed in spam for a real test (an unfamiliar external sender is exactly what mail
 * security gateways are built to be suspicious of). Replaced with MailApp, Apps Script's own
 * built-in mail service -- the same approach already working reliably in this district's MTSS
 * Apps Script project. MailApp sends as the actual signed-in Google Workspace account running
 * this script (Session.getEffectiveUser()), a real trusted internal easdpa.org sender, not an
 * external one -- no API key, no domain verification, no separate account to set up at all.
 * Only real constraint: MailApp's daily send quota (tied to the Workspace account), which is
 * far more than this app's traffic needs.
 *
 * Called from the router as:
 *   sendFollowupEmails_(payload) -- { initials, squad, grade, triggerType, trigDateStr, outcome, attendees: [{name,email}] }
 *                                    or, with payload.singleEvent = true: { ..., eventDate, eventTime, duration, summaryText }
 *   sendSummaryEmail_(payload)   -- { toEmail, subject, summaryText, studentLabel, grade, triggerDate }
 */

// ── SHARED EMAIL LOOK ──────────────────────────────────────────────────────
// One visual shell for every outgoing email, matching the app's own trail palette
// (--trail #2c4a35 / --trail-deep #1a2e1f / --gold #c9a84c / --paper #fffbf5).
// Georgia is used instead of the app's display font (Walter Turncoat) because custom web
// fonts don't render reliably in email clients.
function emailShell_(icon, title, subtitle, bodyHtml) {
  return '<div style="font-family:Georgia,\'Times New Roman\',serif;max-width:560px;margin:0 auto;background:#fffbf5">' +
    '<div style="background:linear-gradient(135deg,#2c4a35 0%,#1a2e1f 100%);padding:26px 24px;border-bottom:3px solid #c9a84c">' +
      '<div style="font-size:26px;margin-bottom:6px">' + icon + '</div>' +
      '<h2 style="margin:0;color:#fff;font-weight:normal;font-size:20px">' + title + '</h2>' +
      (subtitle ? '<p style="margin:6px 0 0;color:rgba(255,255,255,.72);font-size:13px;font-family:Verdana,sans-serif">' + subtitle + '</p>' : '') +
    '</div>' +
    '<div style="padding:22px 24px;background:#fffbf5">' + bodyHtml + '</div>' +
    '<div style="background:#1a2e1f;padding:14px;text-align:center;font-size:11px;color:rgba(255,255,255,.5);font-family:Verdana,sans-serif">Camp Mountaineer &middot; Trail Journal</div>' +
  '</div>';
}

function emailInfoRow_(label, value) {
  return '<p style="margin:0 0 10px;font-family:Verdana,sans-serif;font-size:13px;color:#3a3a3a"><strong style="color:#2c4a35">' + label + ':</strong> ' + value + '</p>';
}

function emailCallout_(title, body, opts) {
  opts = opts || {};
  const bg = opts.bg || '#faf6ec';
  const border = opts.border || '#c9a84c';
  return '<div style="background:' + bg + ';border-left:4px solid ' + border + ';padding:12px 14px;margin:14px 0;border-radius:0 8px 8px 0;font-family:Verdana,sans-serif">' +
    (title ? '<strong style="color:#2c4a35;font-size:13px">' + title + '</strong><br/>' : '') +
    '<span style="font-size:13px;color:#3a3a3a;line-height:1.6">' + body + '</span></div>';
}

function icsFooterNote_(filename) {
  return '<p style="background:#faf6ec;border:1px solid #c9a84c;padding:10px 12px;font-size:12px;font-family:Verdana,sans-serif;border-radius:6px;color:#3a3a3a">' +
    '&#128197; A calendar file (' + filename + ') is attached below -- open it to add this reminder to Outlook.</p>';
}

// Strips the HTML shell down to a readable plain-text fallback. Mail clients that can't (or
// won't) render HTML fall back to this -- and a body-less HTML-only email is itself a mild
// spam signal, so every send below builds both.
function htmlToPlainText_(html) {
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>|<\/div>|<\/h[1-6]>/gi, '\n\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&middot;/g, '-').replace(/&mdash;/g, '--').replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&#\d+;/g, '')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function icsNow_() {
  return Utilities.formatDate(new Date(), 'Etc/UTC', "yyyyMMdd'T'HHmmss'Z'");
}

function icsDate_(ds, h, m) {
  return ds.replace(/-/g, '') + 'T' + String(h).padStart(2, '0') + String(m).padStart(2, '0') + '00';
}

function fmtDate_(ds) {
  const d = new Date(ds + 'T12:00:00');
  return Utilities.formatDate(d, 'America/New_York', 'EEEE, MMMM d, yyyy');
}

function addAndSnap_(ds, days) {
  const d = new Date(ds + 'T12:00:00');
  d.setDate(d.getDate() + days);
  if (d.getDay() === 0) d.setDate(d.getDate() + 1); // Sun -> Mon
  if (d.getDay() === 6) d.setDate(d.getDate() + 2); // Sat -> Mon
  return Utilities.formatDate(d, 'America/New_York', 'yyyy-MM-dd');
}

const GRADE_WINDOWS_ = {
  8: { startH: 8, startM: 15, endH: 8, endM: 45, label: '8:15 - 8:45 AM' },
  7: { startH: 9, startM: 50, endH: 10, endM: 20, label: '9:50 - 10:20 AM' },
};

function buildIcsVevent_(opts) {
  // opts: { uid, startISO, endISO, summary, description, organizerEmail, attendeeEmail, attendeeName }
  return 'BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//EMS Trail Journal//EN\r\nCALSCALE:GREGORIAN\r\nMETHOD:REQUEST\r\n' +
    'BEGIN:VEVENT\r\nUID:' + opts.uid + '\r\nDTSTAMP:' + icsNow_() + '\r\n' +
    'DTSTART;TZID=America/New_York:' + opts.startISO + '\r\nDTEND;TZID=America/New_York:' + opts.endISO + '\r\n' +
    'SUMMARY:' + opts.summary + '\r\nDESCRIPTION:' + opts.description + '\r\nORGANIZER:mailto:' + opts.organizerEmail + '\r\n' +
    'ATTENDEE;CN=' + (opts.attendeeName || opts.attendeeEmail) + ';RSVP=TRUE:mailto:' + opts.attendeeEmail + '\r\n' +
    'STATUS:CONFIRMED\r\nBEGIN:VALARM\r\nTRIGGER:-PT60M\r\nACTION:DISPLAY\r\nDESCRIPTION:Reminder\r\nEND:VALARM\r\nEND:VEVENT\r\nEND:VCALENDAR';
}

// Sends via MailApp (see file header for why). opts: { to, subject, html, icsString, icsFilename }
function sendViaMail_(opts) {
  const mailOptions = {
    to: opts.to,
    subject: opts.subject,
    body: htmlToPlainText_(opts.html),
    htmlBody: opts.html,
    name: 'EMS Trail Journal',
  };
  if (opts.icsString) {
    mailOptions.attachments = [
      Utilities.newBlob(opts.icsString, 'text/calendar; charset=UTF-8; method=REQUEST', opts.icsFilename || 'invite.ics'),
    ];
  }
  MailApp.sendEmail(mailOptions);
}

// ── FOLLOW-UP CHECK-IN SEQUENCE (2-day / 2-week / 2-month) ─────────────────
function sendFollowupEmails_(payload) {
  const ORGANIZER_EMAIL = Session.getEffectiveUser().getEmail();

  // singleEvent: true routes to a genuinely single, custom-dated calendar invite instead of
  // the fixed 2/14/60-day sequence below. This used to be silently ignored -- the frontend's
  // "Schedule Follow-Up Check-In" form (a single date/time/duration picker, sent with
  // singleEvent/eventDate/eventTime/duration) always got the fixed 3-interval sequence
  // instead, because this function never read those fields. Fixed: that form only ever
  // collects ONE date/time, so it should only ever send ONE invite. Also bundles the AI plan
  // summary (payload.summaryText) into that same email, per Matt's request to stop sending
  // the plan summary and the calendar invite as two separate emails.
  if (payload.singleEvent) {
    return sendSingleFollowupEvent_(payload, ORGANIZER_EMAIL);
  }

  const initials = payload.initials;
  const squad = payload.squad;
  const grade = payload.grade;
  const triggerType = payload.triggerType;
  const trigDateStr = payload.trigDateStr;
  const outcome = payload.outcome;
  const attendees = payload.attendees || [];
  const trigLabel = triggerType === 'intervention' ? 'Intervention' : 'Trailback Panel';
  const win = GRADE_WINDOWS_[grade] || GRADE_WINDOWS_[8];

  const GUIDES = {
    2: { title: 'Early Pulse', detail: 'Keep it brief and relational. Ask: How are you doing? Look for re-engagement. Duration: 5-10 minutes.' },
    14: { title: 'First Assessment', detail: 'Grades, attendance, relationships. Is the intervention holding? Duration: 10-20 minutes.' },
    60: { title: 'Long-Term Accountability', detail: 'Pull the data. Has behavior changed? Celebrate growth. Determine if more support needed. Duration: 15-30 minutes.' },
  };

  const intervals = [
    { days: 2, label: '2-Day Check-In', key: 'day2' },
    { days: 14, label: '2-Week Review', key: 'week2' },
    { days: 60, label: '2-Month Follow-Up', key: 'month2' },
  ];

  const results = {};
  const errors = [];

  intervals.forEach(function (interval) {
    const dateStr = addAndSnap_(trigDateStr, interval.days);
    const guide = GUIDES[interval.days];
    const startISO = icsDate_(dateStr, win.startH, win.startM);
    const endISO = icsDate_(dateStr, win.endH, win.endM);
    const subject = '[EMS Follow-Up] ' + initials + ' - ' + interval.label + ' (Grade ' + grade + ', ' + trigLabel + ')';

    let intervalFailed = false;

    attendees.forEach(function (att) {
      try {
        const uid = 'fu-' + interval.key + '-' + new Date().getTime() + '@trailjournal.org';
        const desc = interval.label + ' - ' + trigLabel + '\\nStudent: ' + initials + '\\nGrade: ' + grade +
          '\\nDate: ' + fmtDate_(trigDateStr) + '\\nTime: ' + win.label + '\\nPurpose: ' + guide.title + ' - ' + guide.detail +
          (outcome ? '\\nOutcome: ' + outcome : '');

        const ics = buildIcsVevent_({
          uid: uid, startISO: startISO, endISO: endISO, summary: subject, description: desc,
          organizerEmail: ORGANIZER_EMAIL, attendeeEmail: att.email, attendeeName: att.name,
        });

        const body =
          emailInfoRow_('Student', initials + (squad ? ' &mdash; ' + squad + ' Squad' : '')) +
          emailInfoRow_(trigLabel + ' Date', fmtDate_(trigDateStr)) +
          emailInfoRow_('Check-In Date', fmtDate_(dateStr)) +
          emailInfoRow_('Time', win.label) +
          emailCallout_(guide.title, guide.detail) +
          (outcome ? emailCallout_('Outcome', outcome, { bg: '#f0ebe0', border: '#8a6820' }) : '') +
          icsFooterNote_('followup-' + interval.key + '.ics');

        const html = emailShell_('&#127956;', interval.label, trigLabel + ' Follow-Up &middot; Grade ' + grade, body);

        sendViaMail_({
          to: att.email,
          subject: subject,
          html: html,
          icsString: ics,
          icsFilename: 'followup-' + interval.key + '.ics',
        });
      } catch (e) {
        errors.push(interval.label + ' -> ' + att.email + ': ' + e.message);
        intervalFailed = true;
      }
    });

    results[interval.key] = intervalFailed ? 'failed' : 'sent';
  });

  return { ok: true, results: results, errors: errors };
}

// ── SINGLE CUSTOM FOLLOW-UP EVENT (with the plan summary bundled in) ───────
function sendSingleFollowupEvent_(payload, organizerEmail) {
  const initials = payload.initials;
  const squad = payload.squad;
  const eventDate = payload.eventDate;   // 'yyyy-MM-dd'
  const eventTime = payload.eventTime;   // 'HH:MM', 24-hour
  const duration = Number(payload.duration) || 15; // minutes
  const outcome = payload.outcome;
  const summaryText = payload.summaryText || '';
  const attendees = payload.attendees || [];

  if (!eventDate || !eventTime) throw new Error('eventDate and eventTime are required for a single follow-up event.');

  const [h, m] = eventTime.split(':').map(Number);
  const startISO = icsDate_(eventDate, h, m);
  const endMinutesTotal = h * 60 + m + duration;
  const endH = Math.floor(endMinutesTotal / 60) % 24;
  const endM = endMinutesTotal % 60;
  const endISO = icsDate_(eventDate, endH, endM);

  const timeLabel = formatTimeLabel_(h, m) + ' - ' + formatTimeLabel_(endH, endM);
  const subject = '[EMS Follow-Up] ' + initials + ' - Follow-Up Check-In' + (squad ? ' (' + squad + ' Squad)' : '');

  const escaped = summaryText.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const summaryHtml = escaped
    ? '<div style="margin-top:16px;padding-top:14px;border-top:1px solid #e5ddc8">' +
        escaped.split(/\n\n+/).map(function (para) {
          return '<p style="margin:0 0 12px;font-size:13px;line-height:1.7;color:#3a3a3a;font-family:Verdana,sans-serif;white-space:pre-line">' + para + '</p>';
        }).join('') +
      '</div>'
    : '';

  const results = {};
  const errors = [];
  let anyFailed = false;

  attendees.forEach(function (att) {
    try {
      const uid = 'fu-single-' + new Date().getTime() + '-' + Math.floor(Math.random() * 10000) + '@trailjournal.org';
      const desc = 'Follow-Up Check-In\\nStudent: ' + initials + (squad ? ' (' + squad + ' Squad)' : '') +
        '\\nDate: ' + fmtDate_(eventDate) + '\\nTime: ' + timeLabel + ' (' + duration + ' min)' +
        (outcome ? '\\nContext: ' + outcome : '');

      const ics = buildIcsVevent_({
        uid: uid, startISO: startISO, endISO: endISO, summary: subject, description: desc,
        organizerEmail: organizerEmail, attendeeEmail: att.email, attendeeName: att.name,
      });

      const body =
        emailInfoRow_('Student', initials + (squad ? ' &mdash; ' + squad + ' Squad' : '')) +
        emailInfoRow_('Date', fmtDate_(eventDate)) +
        emailInfoRow_('Time', timeLabel + ' (' + duration + ' min)') +
        (outcome ? emailCallout_('Context', outcome) : '') +
        icsFooterNote_('follow-up-checkin.ics') +
        summaryHtml;

      const html = emailShell_('&#128197;', 'Follow-Up Check-In', initials + (squad ? ' &middot; ' + squad + ' Squad' : ''), body);

      sendViaMail_({
        to: att.email,
        subject: subject,
        html: html,
        icsString: ics,
        icsFilename: 'follow-up-checkin.ics',
      });
    } catch (e) {
      errors.push('Follow-up check-in -> ' + att.email + ': ' + e.message);
      anyFailed = true;
    }
  });

  results.followup = anyFailed ? 'failed' : 'sent';
  return { ok: true, results: results, errors: errors };
}

function formatTimeLabel_(h, m) {
  const ampm = h >= 12 ? 'PM' : 'AM';
  const h12 = ((h + 11) % 12) + 1;
  return h12 + ':' + String(m).padStart(2, '0') + ' ' + ampm;
}

// ── PLAN SUMMARY EMAIL (currently unused by the frontend -- see 04's header) ────
function sendSummaryEmail_(payload) {
  const ORGANIZER_EMAIL = Session.getEffectiveUser().getEmail();

  const toEmail = payload.toEmail;
  const subject = payload.subject || 'Trail Journal Summary';
  const summaryText = payload.summaryText || '';
  const studentLabel = payload.studentLabel || 'Student';
  const grade = payload.grade;
  const triggerDate = payload.triggerDate || Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd');

  if (!toEmail) throw new Error('toEmail is required.');

  const win = GRADE_WINDOWS_[grade] || GRADE_WINDOWS_[8];
  const checkinDate = addAndSnap_(triggerDate, 2);
  const startISO = icsDate_(checkinDate, win.startH, win.startM);
  const endISO = icsDate_(checkinDate, win.endH, win.endM);
  const uid = 'summary-checkin-' + new Date().getTime() + '@trailjournal.org';
  const icsSummary = '[EMS Follow-Up] ' + studentLabel + ' - 2-Day Check-In';
  const desc = '2-Day Check-In\\nStudent: ' + studentLabel + '\\nDate: ' + fmtDate_(checkinDate) + '\\nTime: ' + win.label +
    '\\nPurpose: Early pulse check -- keep it brief and relational.';

  const ics = buildIcsVevent_({
    uid: uid, startISO: startISO, endISO: endISO, summary: icsSummary, description: desc,
    organizerEmail: ORGANIZER_EMAIL, attendeeEmail: toEmail, attendeeName: toEmail,
  });

  // summaryText is plain text built client-side (paragraphs separated by blank lines) --
  // convert to simple HTML paragraphs so it reads cleanly in an email client.
  const escaped = summaryText
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const summaryHtml = escaped.split(/\n\n+/)
    .map(function (para) { return '<p style="margin:0 0 12px;font-size:13px;line-height:1.7;color:#3a3a3a;font-family:Verdana,sans-serif;white-space:pre-line">' + para + '</p>'; })
    .join('');

  const body =
    emailInfoRow_('Student', studentLabel) +
    emailCallout_('2-Day Check-In Reminder', 'Scheduled for ' + fmtDate_(checkinDate) + ', ' + win.label + '.') +
    icsFooterNote_('trail-journal-checkin.ics') +
    '<div style="margin-top:18px;padding-top:16px;border-top:1px solid #e5ddc8">' + summaryHtml + '</div>';

  const html = emailShell_('&#128220;', 'Trail Journal Summary', studentLabel, body);

  sendViaMail_({
    to: toEmail,
    subject: subject,
    html: html,
    icsString: ics,
    icsFilename: 'trail-journal-checkin.ics',
  });

  return { ok: true, sent: true };
}
