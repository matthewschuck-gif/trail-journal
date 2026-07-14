/**
 * Trail Journal -- Apps Script backend
 * File 4 of N: Follow-up check-in emails + Outlook-compatible calendar invites.
 *
 * Faithful port of the live Supabase edge function `send-followup-emails` (v1), pulled
 * directly from the Supabase project. Same three-interval schedule (2-day / 2-week / 2-month),
 * same grade-based time windows, same hand-built .ics attachment so staff on Outlook get a
 * normal calendar invite. Only the transport changed: UrlFetchApp instead of Deno's fetch.
 *
 * NOTE: this still uses Resend to actually send the email (not Gmail/Calendar API), because
 * the school's Google Workspace account doesn't have Gmail enabled -- see the "Follow-up
 * emails" decision from the requirements pass. Resend key lives in Script Properties, never
 * in code.
 *
 * Called from the router as: sendFollowupEmails_(payload)
 * Expected payload shape (matches the original request body exactly):
 * {
 *   initials, squad, grade, triggerType, trigDateStr, outcome,
 *   attendees: [{ name, email }, ...]
 * }
 */

function sendFollowupEmails_(payload) {
  const RESEND_KEY = getProp_('RESEND_API_KEY');
  const FROM_EMAIL = PropertiesService.getScriptProperties().getProperty('FROM_EMAIL') || 'noreply@trailjournal.org';

  const initials = payload.initials;
  const squad = payload.squad;
  const grade = payload.grade;
  const triggerType = payload.triggerType;
  const trigDateStr = payload.trigDateStr;
  const outcome = payload.outcome;
  const attendees = payload.attendees || [];
  const trigLabel = triggerType === 'intervention' ? 'Intervention' : 'Trailback Panel';

  const GRADE_WINDOWS = {
    8: { startH: 8, startM: 15, endH: 8, endM: 45, label: '8:15 - 8:45 AM' },
    7: { startH: 9, startM: 50, endH: 10, endM: 20, label: '9:50 - 10:20 AM' },
  };
  const win = GRADE_WINDOWS[grade] || GRADE_WINDOWS[8];

  const GUIDES = {
    2: { title: 'Early Pulse', detail: 'Keep it brief and relational. Ask: How are you doing? Look for re-engagement. Duration: 5-10 minutes.' },
    14: { title: 'First Assessment', detail: 'Grades, attendance, relationships. Is the intervention holding? Duration: 10-20 minutes.' },
    60: { title: 'Long-Term Accountability', detail: 'Pull the data. Has behavior changed? Celebrate growth. Determine if more support needed. Duration: 15-30 minutes.' },
  };

  function addAndSnap(ds, days) {
    const d = new Date(ds + 'T12:00:00');
    d.setDate(d.getDate() + days);
    if (d.getDay() === 0) d.setDate(d.getDate() + 1); // Sun -> Mon
    if (d.getDay() === 6) d.setDate(d.getDate() + 2); // Sat -> Mon
    return Utilities.formatDate(d, 'America/New_York', 'yyyy-MM-dd');
  }

  function toICSDate(ds, h, m) {
    return ds.replace(/-/g, '') + 'T' + String(h).padStart(2, '0') + String(m).padStart(2, '0') + '00';
  }

  function fmtDate(ds) {
    const d = new Date(ds + 'T12:00:00');
    return Utilities.formatDate(d, 'America/New_York', 'EEEE, MMMM d, yyyy');
  }

  const intervals = [
    { days: 2, label: '2-Day Check-In', key: 'day2' },
    { days: 14, label: '2-Week Review', key: 'week2' },
    { days: 60, label: '2-Month Follow-Up', key: 'month2' },
  ];

  const results = {};
  const errors = [];

  intervals.forEach(function (interval) {
    const dateStr = addAndSnap(trigDateStr, interval.days);
    const guide = GUIDES[interval.days];
    const startISO = toICSDate(dateStr, win.startH, win.startM);
    const endISO = toICSDate(dateStr, win.endH, win.endM);
    const subject = '[EMS Follow-Up] ' + initials + ' - ' + interval.label + ' (Grade ' + grade + ', ' + trigLabel + ')';
    const now = Utilities.formatDate(new Date(), 'Etc/UTC', "yyyyMMdd'T'HHmmss'Z'");

    let intervalFailed = false;

    attendees.forEach(function (att) {
      try {
        const uid = 'fu-' + interval.key + '-' + new Date().getTime() + '@trailjournal.org';
        const desc = interval.label + ' - ' + trigLabel + '\\nStudent: ' + initials + '\\nGrade: ' + grade +
          '\\nDate: ' + fmtDate(trigDateStr) + '\\nTime: ' + win.label + '\\nPurpose: ' + guide.title + ' - ' + guide.detail +
          (outcome ? '\\nOutcome: ' + outcome : '');

        const ics = 'BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//EMS Trail Journal//EN\r\nCALSCALE:GREGORIAN\r\nMETHOD:REQUEST\r\n' +
          'BEGIN:VEVENT\r\nUID:' + uid + '\r\nDTSTAMP:' + now + '\r\n' +
          'DTSTART;TZID=America/New_York:' + startISO + '\r\nDTEND;TZID=America/New_York:' + endISO + '\r\n' +
          'SUMMARY:' + subject + '\r\nDESCRIPTION:' + desc + '\r\nORGANIZER:mailto:' + FROM_EMAIL + '\r\n' +
          'ATTENDEE;CN=' + (att.name || att.email) + ';RSVP=TRUE:mailto:' + att.email + '\r\n' +
          'STATUS:CONFIRMED\r\nBEGIN:VALARM\r\nTRIGGER:-PT60M\r\nACTION:DISPLAY\r\nDESCRIPTION:Reminder\r\nEND:VALARM\r\nEND:VEVENT\r\nEND:VCALENDAR';

        const html = '<div style="font-family:Georgia,serif;max-width:560px;margin:0 auto">' +
          '<div style="background:#490e6f;padding:24px;color:#ffe100"><h2 style="margin:0">' + interval.label + '</h2>' +
          '<p style="margin:4px 0 0;color:rgba(255,255,255,.75);font-size:13px">' + trigLabel + ' Follow-Up - Grade ' + grade + '</p></div>' +
          '<div style="padding:20px;background:#fff"><p><strong>Student:</strong> ' + initials + (squad ? ' - ' + squad + ' Squad' : '') + '</p>' +
          '<p><strong>' + trigLabel + ' Date:</strong> ' + fmtDate(trigDateStr) + '</p>' +
          '<p><strong>Check-In Date:</strong> ' + fmtDate(dateStr) + '</p>' +
          '<p><strong>Time:</strong> ' + win.label + '</p>' +
          '<div style="background:#faf7f2;border-left:4px solid #490e6f;padding:12px;margin:16px 0"><strong>' + guide.title + '</strong><br/>' +
          '<span style="font-size:13px">' + guide.detail + '</span></div>' +
          (outcome ? '<div style="background:#f0ebf8;padding:12px;border-radius:8px;margin-bottom:12px"><strong>Outcome:</strong> ' + outcome + '</div>' : '') +
          '<p style="background:#fdf5df;border:1px solid #c9a84c;padding:10px;font-size:13px;border-radius:6px">A .ics file is attached - open it to add to your Outlook calendar.</p></div>' +
          '<div style="background:#1a2e1f;padding:12px;text-align:center;font-size:11px;color:rgba(255,255,255,.5)">EMS Mountaineers - Trail Journal</div></div>';

        const resendPayload = {
          from: 'EMS Trail Journal <' + FROM_EMAIL + '>',
          to: [att.email],
          subject: subject,
          html: html,
          attachments: [{ filename: 'followup-' + interval.key + '.ics', content: Utilities.base64Encode(ics) }],
        };

        const r = UrlFetchApp.fetch('https://api.resend.com/emails', {
          method: 'post',
          contentType: 'application/json',
          headers: { Authorization: 'Bearer ' + RESEND_KEY },
          payload: JSON.stringify(resendPayload),
          muteHttpExceptions: true,
        });

        if (r.getResponseCode() >= 300) {
          throw new Error('Resend ' + r.getResponseCode() + ': ' + r.getContentText());
        }
      } catch (e) {
        errors.push(interval.label + ' -> ' + att.email + ': ' + e.message);
        intervalFailed = true;
      }
    });

    results[interval.key] = intervalFailed ? 'failed' : 'sent';
  });

  return { ok: true, results: results, errors: errors };
}
