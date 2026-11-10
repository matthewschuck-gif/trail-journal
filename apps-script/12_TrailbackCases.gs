/**
 * Trail Journal -- Apps Script backend
 * File 12: Trailback Panel cases, from the first suggestion to the close.
 *
 * Status, in order:
 *   suggested -> offered -> consented -> tj_done -> scheduled -> heard_approved -> repair_in_progress
 *   -> closed_completed | closed_reverted.   declined can happen any time before the panel.
 *
 * One case per student per consequence (panel_cases tab). Every change goes through
 * updateTrailbackCase_ (one step at a time; out-of-order moves are refused) and is written to
 * the case's history with who did it.
 *
 * The detention itself lives on the Accountability & Intervention Log's Master Log (another
 * Sheet, ACCOUNTABILITY_SHEET_ID), matched by Entry ID (column A) and 800# (column AH):
 *   consent          -> Status "Held (Trailback)"     (the student doesn't serve it while waiting)
 *   close completed  -> Status "Resolved (Trailback)" + an anonymous Clash point request
 *   close reverted   -> Status "Scheduled" + a note to pick a new date (never booked for them)
 *   declined (after consent) -> Status "Scheduled" + the same note
 *
 * Emails (OFF until Settings TRAILBACK_EMAILS = ON; menu items preview them with nothing sent):
 *   Tuesday about 3:00 p.m. -- the case brief for Wednesday's panel, to the adults on the Panel
 *     Members tab (rows with Role "Student" are skipped).
 *   Weekday mornings -- overdue 3/10/30 school-day check-ins, to TRAILBACK_COORDINATOR_EMAIL.
 *   A case closed as reverted -- the same coordinator, to pick a new detention date.
 *
 * Other tools (server to server, a shared key from Script Properties TRAILBACK_API_KEY):
 *   createTrailbackCase      -- the Consequence Entry Form (Accountability system), status offered.
 *   trailbackCasesByStudent  -- Climb-er: status, dates and outcome only, no notes.
 */

const CASE_STATUSES_ = ['suggested', 'offered', 'consented', 'tj_done', 'scheduled', 'heard_approved',
  'repair_in_progress', 'closed_completed', 'closed_reverted', 'declined'];
const CASE_LABELS_ = {
  suggested: 'Suggested', offered: 'Offered', consented: 'Consent received', tj_done: 'Trail Journal done',
  scheduled: 'Scheduled', heard_approved: 'Heard and approved', repair_in_progress: 'Repair in progress',
  closed_completed: 'Closed: completed', closed_reverted: 'Closed: reverted', declined: 'Declined',
};
const PRE_PANEL_ = ['suggested', 'offered', 'consented', 'tj_done', 'scheduled'];
const CASE_SOURCES_ = ['admin_conference', 'teacher_referral', 'climber'];
const CASE_DATE_FIELDS_ = ['consequence_original_date', 'scheduled_date', 'checkin_3', 'checkin_10', 'checkin_30'];
const CASE_TIME_FIELDS_ = ['created_at', 'offered_at', 'consent_at', 'fixit_meeting_at', 'closed_at', 'updated_at'];
const ML_TAB_ = 'Master Log';
const ML_STATUS_COL_ = 16;   // P
const ML_NOTES_COL_ = 17;    // Q
const ML_ID_INDEX_ = 0;      // A  Entry ID
const ML_800_INDEX_ = 33;    // AH Student 800#
const PANEL_MEMBERS_TAB_ = 'Panel Members';
const TEST_CASE_PREFIX_ = 'TB-TEST-';

// ── READING CASES ───────────────────────────────────────────────────
function digits800_(v) { return String(v == null ? '' : v).replace(/\D/g, ''); }
function todayYmd_() { return ymdOf_(new Date()); }

/** A case as the pages and other tools see it: dates as yyyy-MM-dd, times as ISO text. */
function caseOut_(row) {
  const c = Object.assign({}, row);
  CASE_DATE_FIELDS_.forEach(function (k) { c[k] = ymdOf_(c[k]); });
  CASE_TIME_FIELDS_.forEach(function (k) { if (c[k] instanceof Date) c[k] = c[k].toISOString(); });
  ['checkin_notes', 'history', 'referral_json'].forEach(function (k) {
    if (typeof c[k] === 'string' && c[k]) { try { c[k] = JSON.parse(c[k]); } catch (e) { /* leave as text */ } }
  });
  if (!c.checkin_notes || typeof c.checkin_notes !== 'object') c.checkin_notes = {};
  if (!Array.isArray(c.history)) c.history = [];
  c.status_label = CASE_LABELS_[c.status] || c.status;
  return c;
}

function allCases_() {
  ensureSchemaTab_(TABS.PANEL_CASES);
  return genericQuery_(TABS.PANEL_CASES, {}).map(caseOut_);
}

function getCase_(caseId) {
  ensureSchemaTab_(TABS.PANEL_CASES);
  const r = genericQuery_(TABS.PANEL_CASES, { case_id: caseId }, 1)[0];
  if (!r) throw new Error('No Trailback case ' + caseId);
  return caseOut_(r);
}

function saveCase_(c, patch, by, from, note) {
  const history = (c.history || []).concat([{ at: nowIso_(), by: by || '', from: from || c.status, to: patch.status || c.status, note: note || '' }]);
  const full = Object.assign({}, patch, { history: JSON.stringify(history), updated_by: by || '', updated_at: nowIso_() });
  if (full.checkin_notes && typeof full.checkin_notes !== 'string') full.checkin_notes = JSON.stringify(full.checkin_notes);
  genericUpdate_(TABS.PANEL_CASES, { case_id: c.case_id }, full);
  return getCase_(c.case_id);
}

function whoIs_(body) { return body._staffEmail || body._caller || 'staff password'; }

// ── CREATING CASES ──────────────────────────────────────────────────
/**
 * fields: { student_id, student_name, grade, squad, behavior_type, source, referring_staff,
 *   consequence_type, consequence_original_date, master_log_row, incident_summary,
 *   behavior_pattern, referral_json, status ('suggested' | 'offered') }
 * The same student + Master Log row with a case still open gives back that case.
 */
function createCase_(fields, by) {
  const f = fields || {};
  const sid = digits800_(f.student_id);
  if (sid && !/^80\d{5}$/.test(sid)) throw new Error('The 800# should be 7 digits starting with 80 (got ' + f.student_id + ').');
  if (!sid && !String(f.student_name || '').trim()) throw new Error('A case needs the student\'s 800# or name.');
  const status = f.status === 'offered' ? 'offered' : 'suggested';
  const source = CASE_SOURCES_.indexOf(f.source) > -1 ? f.source : 'admin_conference';
  ensureSchemaTab_(TABS.PANEL_CASES);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (sid) {
      const open = genericQuery_(TABS.PANEL_CASES, { student_id: sid }).map(caseOut_).filter(function (c) {
        return c.status.indexOf('closed') !== 0 && c.status !== 'declined' &&
          (!f.master_log_row || !c.master_log_row || String(c.master_log_row) === String(f.master_log_row));
      })[0];
      if (open) return Object.assign(open, { existing: true });
    }
    const now = nowIso_();
    const row = {
      case_id: f.case_id || ('TB-' + Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase()),
      student_id: sid, student_name: String(f.student_name || '').trim(), grade: f.grade || '',
      squad: String(f.squad || '').trim(), behavior_type: f.behavior_type || '', source: source,
      referring_staff: f.referring_staff || '', consequence_type: f.consequence_type || '',
      consequence_original_date: ymdOf_(f.consequence_original_date) || '', master_log_row: f.master_log_row || '',
      status: status, offered_at: status === 'offered' ? now : '',
      incident_summary: f.incident_summary || '', behavior_pattern: f.behavior_pattern || '',
      referral_json: f.referral_json ? JSON.stringify(f.referral_json) : '',
      checkin_notes: '{}', history: JSON.stringify([{ at: now, by: by || '', from: '', to: status, note: 'Case created (' + source + ')' }]),
      updated_by: by || '', updated_at: now,
    };
    genericInsert_(TABS.PANEL_CASES, row);
    return getCase_(row.case_id);
  } finally {
    lock.releaseLock();
  }
}

/** Staff page or another tool: body.payload = the case fields. */
function createTrailbackCase_(body) {
  return createCase_(body.payload || {}, whoIs_(body));
}

// ── MOVING A CASE ───────────────────────────────────────────────────
const EDITABLE_CASE_FIELDS_ = ['student_id', 'student_name', 'grade', 'squad', 'behavior_type', 'referring_staff',
  'consequence_type', 'consequence_original_date', 'master_log_row', 'checkin_person', 'completion_criteria',
  'incident_summary', 'behavior_pattern'];

function isWednesday_(ymd) {
  const p = ymd.split('-').map(Number);
  return new Date(Date.UTC(p[0], p[1] - 1, p[2], 12)).getUTCDay() === 3;
}

/**
 * body.payload: { case_id, op, ... }. ops and where they may start:
 *   offer (suggested) · consent (suggested, offered) · link_tj {trail_journal_id} (consented; sets tj_done)
 *   schedule {date: a Wednesday} (consented, tj_done, scheduled) · checkin {n: 3|10|30, note}
 *   fixit {date?, note} · close_completed {reason?} · close_reverted {reason}
 *   (repair_in_progress, or heard_approved for reverted) · decline {reason} (before the panel)
 *   edit {fields} (any open case) · clear_flag (any)
 * Returns { case, warnings }.
 */
function updateTrailbackCase_(body) {
  const p = body.payload || {};
  const by = whoIs_(body);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const c = getCase_(p.case_id);
    const warnings = [];
    const need = function (allowed, what) {
      if (allowed.indexOf(c.status) === -1) throw new Error('Can\'t ' + what + ' a case that is "' + c.status_label + '".');
    };
    const now = nowIso_();
    let patch = {}, note = p.note || '';

    switch (p.op) {
      case 'offer':
        need(['suggested'], 'offer');
        patch = { status: 'offered', offered_at: now };
        break;
      case 'consent':
        need(['suggested', 'offered'], 'record consent for');
        patch = { status: 'consented', consent_at: now, offered_at: c.offered_at || now };
        warnings.push.apply(warnings, masterLogSet_(c, 'Held (Trailback)', 'Trailback consent ' + todayYmd_() + ': held for the panel'));
        break;
      case 'link_tj':
        if (!p.trail_journal_id) throw new Error('trail_journal_id is required.');
        if (c.status.indexOf('closed') === 0 || c.status === 'declined') throw new Error('This case is closed.');
        patch = { trail_journal_id: String(p.trail_journal_id) };
        if (c.status === 'consented') patch.status = 'tj_done';
        break;
      case 'schedule': {
        need(['consented', 'tj_done', 'scheduled'], 'schedule');
        const d = ymdOf_(p.date);
        if (!d || !isWednesday_(d)) throw new Error('Panels are on Wednesdays. Pick a Wednesday.');
        if (d < todayYmd_()) throw new Error('That Wednesday has already passed.');
        if (noSchoolDates_()[d]) throw new Error(d + ' is on the No School Dates tab.');
        patch = { status: 'scheduled', scheduled_date: d };
        note = note || ('Panel ' + d);
        break;
      }
      case 'checkin': {
        need(['repair_in_progress'], 'log a check-in for');
        const n = Number(p.n);
        if (CHECKIN_SCHOOL_DAYS_.indexOf(n) === -1) throw new Error('Which check-in: 3, 10 or 30?');
        const notes = Object.assign({}, c.checkin_notes);
        const due = c['checkin_' + n];
        notes[n] = { at: now, by: by, note: String(p.note || ''), on_time: !due || todayYmd_() <= due };
        patch = { checkin_notes: notes };
        note = n + '-school-day check-in logged';
        break;
      }
      case 'fixit':
        need(['repair_in_progress', 'heard_approved'], 'log a fix-it meeting for');
        patch = { fixit_meeting_at: p.date ? ymdOf_(p.date) || String(p.date) : now };
        note = 'Fix-it meeting' + (p.note ? ': ' + p.note : '');
        break;
      case 'close_completed': {
        need(['repair_in_progress'], 'close');
        patch = { status: 'closed_completed', closed_at: now, closed_reason: p.reason || 'Plan completed' };
        warnings.push.apply(warnings, masterLogSet_(c, 'Resolved (Trailback)', 'Trailback plan completed ' + todayYmd_() + ': detention cleared'));
        if (c.squad) {
          try { patch.clash_request_id = requestClashPoint_({ squad: c.squad, caseId: c.case_id, requestedBy: by }).id; }
          catch (e) { warnings.push('No Clash point request: ' + e.message); }
        } else {
          warnings.push('No squad on this case, so no Clash point was requested.');
        }
        break;
      }
      case 'close_reverted':
        need(['repair_in_progress', 'heard_approved'], 'revert');
        if (!String(p.reason || '').trim()) throw new Error('Say why the case reverted.');
        patch = { status: 'closed_reverted', closed_at: now, closed_reason: p.reason, admin_flag: 'Pick a new detention date on the Master Log' };
        warnings.push.apply(warnings, masterLogSet_(c, 'Scheduled', 'Trailback reverted ' + todayYmd_() + ': pick a new date'));
        warnings.push.apply(warnings, notifyReverted_(c, p.reason));
        break;
      case 'decline': {
        need(PRE_PANEL_, 'decline');
        patch = { status: 'declined', closed_at: now, closed_reason: p.reason || 'Declined' };
        if (c.consent_at) {
          patch.admin_flag = 'Pick a new detention date on the Master Log';
          warnings.push.apply(warnings, masterLogSet_(c, 'Scheduled', 'Trailback declined ' + todayYmd_() + ': pick a new date'));
        }
        break;
      }
      case 'edit': {
        if (c.status.indexOf('closed') === 0 || c.status === 'declined') throw new Error('This case is closed.');
        const f = p.fields || {};
        EDITABLE_CASE_FIELDS_.forEach(function (k) { if (Object.prototype.hasOwnProperty.call(f, k)) patch[k] = f[k]; });
        if (patch.student_id !== undefined) patch.student_id = digits800_(patch.student_id);
        note = 'Edited: ' + Object.keys(patch).join(', ');
        break;
      }
      case 'clear_flag':
        patch = { admin_flag: '' };
        note = 'Flag cleared';
        break;
      default:
        throw new Error('Unknown case step: ' + p.op);
    }
    return { case: saveCase_(c, patch, by, c.status, note), warnings: warnings };
  } finally {
    lock.releaseLock();
  }
}

// ── THE PANEL SESSION ↔ ITS CASE (called from 11_Trailback.gs) ───────
/** At approval. A panel started without a queued case gets one here. Returns the case_id. */
function casePanelApproved_(record, approval) {
  let c;
  if (record.case_id) {
    c = getCase_(record.case_id);
    if (c.status.indexOf('closed') === 0 || c.status === 'declined') throw new Error('Case ' + c.case_id + ' is ' + c.status_label + '.');
  } else {
    c = createCase_({ student_id: record.student_id, student_name: record.student_name, grade: record.grade,
      squad: record.clash_squad, referring_staff: record.referring_staff, incident_summary: record.incident_summary,
      behavior_pattern: record.behavior_pattern, source: 'admin_conference' }, approval.approved_by);
    if (!c.existing) c = getCase_(c.case_id);
  }
  if (c.status === 'heard_approved' || c.status === 'repair_in_progress') return c.case_id;
  saveCase_(c, { status: 'heard_approved', panel_session_id: record.session_id, scheduled_date: c.scheduled_date || ymdOf_(record.panel_date),
    trail_journal_id: c.trail_journal_id || record.trail_journal_id || '' }, approval.approved_by, c.status, 'Plan approved at the panel');
  return c.case_id;
}

/** At submit: repair starts; the check-in dates, person and criteria go onto the case. */
function casePanelSubmitted_(caseId, record, by) {
  if (!caseId) return;
  const c = getCase_(caseId);
  if (c.status !== 'heard_approved' && c.status !== 'repair_in_progress') return;
  const d = record.followup_dates || {};
  saveCase_(c, { status: 'repair_in_progress', checkin_3: d.checkin_3 || '', checkin_10: d.checkin_10 || '', checkin_30: d.checkin_30 || '',
    checkin_person: record.followup_person || '', completion_criteria: record.fu_criteria || '' }, by, c.status, 'Panel record submitted');
}

// ── TRAIL JOURNAL AUTO-LINK ─────────────────────────────────────────
/** A consented case whose student finished a Trail Journal since consent: link it, status tj_done. */
function autoLinkTrailJournals_() {
  let linked = 0;
  allCases_().filter(function (c) { return c.status === 'consented' && c.student_id && !c.trail_journal_id; }).forEach(function (c) {
    const since = String(c.consent_at || '');
    const done = genericQuery_(TABS.REFLECTIONS, { student_id: c.student_id }).filter(function (r) {
      const at = r.created_at instanceof Date ? r.created_at.toISOString() : String(r.created_at || '');
      return String(r.status || 'submitted') === 'submitted' && at >= since;
    }).sort(function (a, b) { return String(b.created_at) < String(a.created_at) ? -1 : 1; })[0];
    if (!done) return;
    saveCase_(c, { status: 'tj_done', trail_journal_id: done.session_id }, 'auto (Trail Journal finished)', c.status, 'Trail Journal ' + done.session_id);
    linked++;
  });
  return linked;
}

// ── THE BOARD (Admin) ───────────────────────────────────────────────
function nextWednesday_(fromYmd) {
  const p = (fromYmd || todayYmd_()).split('-').map(Number);
  const d = new Date(Date.UTC(p[0], p[1] - 1, p[2], 12));
  while (d.getUTCDay() !== 3) d.setUTCDate(d.getUTCDate() + 1);
  return d.toISOString().slice(0, 10);
}

function weekOf_(ymd) {
  const p = ymd.split('-').map(Number);
  const d = new Date(Date.UTC(p[0], p[1] - 1, p[2], 12));
  d.setUTCDate(d.getUTCDate() - ((d.getUTCDay() + 6) % 7)); // Monday
  const start = d.toISOString().slice(0, 10);
  d.setUTCDate(d.getUTCDate() + 6);
  return { start: start, end: d.toISOString().slice(0, 10) };
}

/** Counts for the week holding `ymd`: offered, accepted, held, completed, reverted, check-ins on time. */
function weeklyCounts_(cases, ymd) {
  const w = weekOf_(ymd);
  // Timestamps are saved in UTC (ISO text); read them as a local school day, so an evening
  // entry counts on the day it happened.
  const inWeek = function (v) {
    const s = v instanceof Date ? v : String(v || '');
    const d = typeof s === 'string' && s.indexOf('T') > -1 ? ymdOf_(new Date(s)) : ymdOf_(s);
    return d && d >= w.start && d <= w.end;
  };
  const out = { week_start: w.start, week_end: w.end, offered: 0, accepted: 0, held: 0, completed: 0, reverted: 0, checkins_due: 0, checkins_on_time: 0 };
  const today = todayYmd_();
  cases.forEach(function (c) {
    if (inWeek(c.offered_at)) out.offered++;
    if (inWeek(c.consent_at)) out.accepted++;
    if ((c.history || []).some(function (h) { return h.to === 'heard_approved' && inWeek(h.at); })) out.held++;
    if (c.status === 'closed_completed' && inWeek(c.closed_at)) out.completed++;
    if (c.status === 'closed_reverted' && inWeek(c.closed_at)) out.reverted++;
    CHECKIN_SCHOOL_DAYS_.forEach(function (n) {
      const due = c['checkin_' + n];
      if (!due || !inWeek(due) || due > today) return;
      out.checkins_due++;
      const logged = c.checkin_notes[n];
      if (logged && logged.on_time) out.checkins_on_time++;
    });
  });
  return out;
}

/** Staff: everything the Admin tab and the Panel queue need. body.payload.week = any date in the week. */
function trailbackCases_(body) {
  let linked = 0;
  try { linked = autoLinkTrailJournals_(); } catch (e) { console.error('autoLinkTrailJournals_: ' + e); }
  const cases = allCases_();
  const today = todayYmd_();
  const wed = nextWednesday_(today);
  return {
    today: today, wednesday: wed, linked: linked,
    cases: cases.sort(function (a, b) { return String(b.updated_at) < String(a.updated_at) ? -1 : 1; }),
    docket: cases.filter(function (c) { return c.status === 'scheduled' && c.scheduled_date === wed; }),
    counts: weeklyCounts_(cases, ymdOf_((body.payload && body.payload.week) || today) || today),
    statuses: CASE_STATUSES_.map(function (s) { return { key: s, label: CASE_LABELS_[s] }; }),
  };
}

function trailbackCase_(body) { return getCase_((body.payload || {}).case_id); }

/** Climb-er: status, dates and outcome only. No notes, names or plan. */
function trailbackCasesByStudent_(body) {
  const sid = digits800_((body.payload || {}).student_id);
  if (!/^80\d{5}$/.test(sid)) throw new Error('student_id (800#) is required.');
  ensureSchemaTab_(TABS.PANEL_CASES);
  return genericQuery_(TABS.PANEL_CASES, { student_id: sid }).map(caseOut_).map(function (c) {
    return {
      case_id: c.case_id, status: c.status, status_label: c.status_label, created_at: c.created_at,
      offered_at: c.offered_at, consent_at: c.consent_at, scheduled_date: c.scheduled_date,
      checkin_3: c.checkin_3, checkin_10: c.checkin_10, checkin_30: c.checkin_30, closed_at: c.closed_at,
      outcome: c.status === 'closed_completed' ? 'completed' : c.status === 'closed_reverted' ? 'reverted'
        : c.status === 'declined' ? 'declined' : 'open',
    };
  });
}

// ── MASTER LOG (Accountability & Intervention Log) ──────────────────
/** Sets the Status of the case's Master Log row and adds a note. Returns warnings (never throws). */
function masterLogSet_(c, status, note) {
  if (!c.master_log_row) return c.source === 'admin_conference' && c.consequence_type ? ['This case has no Master Log Entry ID, so the Master Log was not changed.'] : [];
  const id = webSetting_('ACCOUNTABILITY_SHEET_ID');
  if (!id) return ['ACCOUNTABILITY_SHEET_ID is blank on the Settings tab, so the Master Log was not changed.'];
  try {
    const sh = SpreadsheetApp.openById(id).getSheetByName(ML_TAB_);
    if (!sh) return ['The Accountability Sheet has no "' + ML_TAB_ + '" tab.'];
    const vals = sh.getDataRange().getDisplayValues();
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][ML_ID_INDEX_]).trim() !== String(c.master_log_row).trim()) continue;
      const rowId = digits800_(vals[i][ML_800_INDEX_]);
      if (rowId && c.student_id && rowId !== c.student_id) {
        return ['Master Log row ' + c.master_log_row + ' is for a different 800#, so it was not changed. Check the Entry ID on the case.'];
      }
      sh.getRange(i + 1, ML_STATUS_COL_).setValue(status);
      const old = String(vals[i][ML_NOTES_COL_ - 1] || '');
      sh.getRange(i + 1, ML_NOTES_COL_).setValue(old ? old + ' | ' + note : note);
      return [];
    }
    return ['No Master Log row with Entry ID ' + c.master_log_row + '. Set its Status to "' + status + '" by hand.'];
  } catch (e) {
    return ['Could not open the Accountability Sheet (' + e.message + '). Set the Master Log Status to "' + status + '" by hand.'];
  }
}

// ── EMAILS ──────────────────────────────────────────────────────────
function trailbackEmailsOn_() { return String(webSetting_('TRAILBACK_EMAILS')).toUpperCase() === 'ON'; }
function coordinatorEmails_() { return String(webSetting_('TRAILBACK_COORDINATOR_EMAIL') || '').split(',').map(function (s) { return s.trim(); }).filter(String).join(','); }

/** Adults on the Panel Members tab (Name | Email | Role | Active). Role "Student" rows are skipped. */
function panelMemberEmails_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PANEL_MEMBERS_TAB_);
  if (!sh) return '';
  const rows = sh.getDataRange().getValues();
  const h = rows[0].map(function (x) { return String(x).trim().toLowerCase(); });
  const col = function (name) { return h.indexOf(name); };
  return rows.slice(1).filter(function (r) {
    const email = String(r[col('email')] || '').trim();
    const role = String(r[col('role')] || '').toLowerCase();
    const active = col('active') < 0 ? '' : String(r[col('active')]).toLowerCase();
    return email && role.indexOf('student') === -1 && active !== 'false' && active !== 'no';
  }).map(function (r) { return String(r[col('email')]).trim(); }).join(',');
}

/** Sends (when TRAILBACK_EMAILS is ON and not a preview) and records the result. */
function sendTrailbackMail_(to, subject, html, preview) {
  const result = { to: to, subject: subject, sent: false, at: nowIso_() };
  if (preview) result.why = 'Preview: nothing sent.';
  else if (!trailbackEmailsOn_()) result.why = 'TRAILBACK_EMAILS is OFF on the Settings tab.';
  else if (!to) result.why = 'No recipients.';
  else {
    MailApp.sendEmail({ to: to, subject: subject, htmlBody: html, body: htmlToPlainText_(html), name: 'EMS Trailback Panel' });
    result.sent = true;
  }
  try { PropertiesService.getScriptProperties().setProperty('trailback_last_email', JSON.stringify(result)); } catch (e) { /* tests */ }
  return Object.assign({ html: html }, result);
}

function briefHtml_(wed, cases) {
  const deadline = webSetting_('TRAILBACK_CONFLICT_DEADLINE') || 'Wednesday morning, before the panel';
  const body = cases.map(function (c) {
    // Each case card wears the student's squad color, like the journal's emails (squadTheme_).
    const t = squadTheme_(c.squad);
    const squadLine = t === DEFAULT_THEME_ ? '' : ' &middot; ' + t.emoji + ' ' + t.label + ' squad';
    return '<div style="border:1.5px solid ' + t.primary + ';border-left:6px solid ' + t.primary + ';border-radius:8px;padding:12px 14px;margin:0 0 12px;font-family:Verdana,sans-serif">' +
      emailInfoRow_('Student', escHtml_(c.student_name || '—') + (c.grade ? ' &middot; Grade ' + escHtml_(c.grade) : '') + squadLine, t) +
      emailInfoRow_('Behavior', escHtml_(c.behavior_type || '—')) +
      emailInfoRow_('Incident', escHtml_(c.incident_summary || '—')) +
      emailInfoRow_('Pattern', escHtml_(c.behavior_pattern || '—')) +
      '</div>';
  }).join('') +
    emailCallout_('Conflict of interest?', 'If you know a student on this list, or the people involved, in a way that makes it hard to be fair, tell Kelly by <strong>' + escHtml_(deadline) + '</strong>. Someone else will sit in for that case.') +
    '<p style="font-family:Verdana,sans-serif;font-size:12px;color:#666">Please keep these details to the panel. This email went to the adult panel members only.</p>';
  return emailShell_('&#9978;', 'Trailback Panel Brief', 'Wednesday ' + escHtml_(wed) + ' &middot; ' + cases.length + ' case' + (cases.length === 1 ? '' : 's'), body);
}

/** The Tuesday brief for the coming Wednesday. preview = build it, send nothing. */
function trailbackTuesdayBrief_(preview) {
  const wed = nextWednesday_(todayYmd_());
  if (noSchoolDates_()[wed]) return { skipped: 'No school on ' + wed };
  const cases = allCases_().filter(function (c) { return c.status === 'scheduled' && c.scheduled_date === wed; });
  if (!cases.length && !preview) return { skipped: 'No cases on ' + wed };
  return sendTrailbackMail_(panelMemberEmails_(), 'Trailback Panel brief: Wednesday ' + wed, briefHtml_(wed, cases), preview);
}

/** Repair-in-progress cases with a 3/10/30 check-in past due and not logged. */
function overdueCheckins_() {
  const today = todayYmd_();
  const out = [];
  allCases_().filter(function (c) { return c.status === 'repair_in_progress'; }).forEach(function (c) {
    CHECKIN_SCHOOL_DAYS_.forEach(function (n) {
      const due = c['checkin_' + n];
      if (due && due < today && !c.checkin_notes[n]) out.push({ c: c, n: n, due: due });
    });
  });
  return out;
}

function trailbackOverdueEmail_(preview) {
  const list = overdueCheckins_();
  if (!list.length && !preview) return { skipped: 'Nothing overdue' };
  const body = (list.length ? list.map(function (o) {
    const t = squadTheme_(o.c.squad);
    return emailInfoRow_((t === DEFAULT_THEME_ ? '' : t.emoji + ' ') + escHtml_(o.c.student_name || o.c.student_id || o.c.case_id), o.n + '-school-day check-in was due ' + o.due +
      (o.c.checkin_person ? ' &middot; check-in person: ' + escHtml_(o.c.checkin_person) : ''));
  }).join('') : '<p style="font-family:Verdana,sans-serif">Nothing is overdue today.</p>') +
    emailCallout_('Log it', 'Open Trail Journal Admin &rsaquo; Trailback Cases and use <strong>Log check-in</strong> on each case.');
  return sendTrailbackMail_(coordinatorEmails_(), 'Trailback: ' + list.length + ' overdue check-in' + (list.length === 1 ? '' : 's'),
    emailShell_('&#9200;', 'Overdue Trailback Check-Ins', todayYmd_(), body), preview);
}

function notifyReverted_(c, reason) {
  const t = squadTheme_(c.squad); // the student's squad colors, like the journal's emails
  const html = emailShell_('&#8617;', 'Trailback case reverted', escHtml_(c.student_name || c.student_id || c.case_id) +
      (t === DEFAULT_THEME_ ? '' : ' &middot; ' + t.label + ' squad'),
    emailInfoRow_('Reason', escHtml_(reason), t) +
    emailInfoRow_('Master Log Entry ID', escHtml_(c.master_log_row || 'none on the case'), t) +
    emailCallout_('Pick a new date', 'The Master Log row is back to <strong>Scheduled</strong>. Pick a new detention date there; nothing was booked automatically.', null, t), t);
  const r = sendTrailbackMail_(coordinatorEmails_(), 'Trailback reverted: pick a new detention date', html, false);
  return r.sent ? [] : ['Reverted email not sent: ' + r.why];
}

// ── TRIGGERS (menu) ─────────────────────────────────────────────────
function trailbackTuesdayBrief() { trailbackTuesdayBrief_(false); }
function trailbackDailyCheck() {
  const dow = new Date().getDay();
  if (dow === 0 || dow === 6 || noSchoolDates_()[todayYmd_()]) return;
  try { autoLinkTrailJournals_(); } catch (e) { console.error(e); }
  trailbackOverdueEmail_(false);
}

function setupTrailbackTriggers() {
  ['trailbackTuesdayBrief', 'trailbackDailyCheck'].forEach(function (fn) {
    ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === fn) ScriptApp.deleteTrigger(t); });
  });
  ScriptApp.newTrigger('trailbackTuesdayBrief').timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(15).nearMinute(0).create();
  ScriptApp.newTrigger('trailbackDailyCheck').timeBased().everyDays(1).atHour(7).create();
  safeAlert_('Trailback timers are on: the Tuesday brief (about 3:00 p.m.) and the weekday overdue check (about 7 a.m.).\n\n' +
    'Emails only go out when Settings TRAILBACK_EMAILS is ON (it is ' + (trailbackEmailsOn_() ? 'ON' : 'OFF') + ').');
}

function previewDialog_(r, title) {
  const head = r.skipped ? '<p><b>' + escHtml_(r.skipped) + '</b></p>' :
    '<p style="font-family:Verdana;font-size:12px"><b>To:</b> ' + escHtml_(r.to || '(nobody yet)') + '<br><b>Subject:</b> ' + escHtml_(r.subject) + '<br>' + escHtml_(r.why || '') + '</p>';
  try {
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(head + (r.html || '')).setWidth(640).setHeight(620), title);
  } catch (e) { Logger.log(head); }
}
function previewTuesdayBrief() { previewDialog_(trailbackTuesdayBrief_(true), 'Tuesday brief preview (nothing sent)'); }
function previewOverdueEmail() { previewDialog_(trailbackOverdueEmail_(true), 'Overdue check-ins preview (nothing sent)'); }

// ── TEACHER REFERRAL FORM ───────────────────────────────────────────
function connectReferralForm() {
  const id = webSetting_('TRAILBACK_REFERRAL_FORM_ID');
  if (!id) { safeAlert_('Put the referral Form ID in Settings TRAILBACK_REFERRAL_FORM_ID first.'); return; }
  FormApp.openById(id); // fails here (with the reason) if this account can't open the Form
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'onTrailbackReferral') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onTrailbackReferral').forForm(id).onFormSubmit().create();
  safeAlert_('The Trailback Panel Staff Referral form is connected. Each new referral becomes a "Suggested" case on Admin › Trailback Cases.');
}

/** answers: { question title: answer }. Questions are matched by words in their titles. */
function referralToCase_(answers, respondentEmail) {
  const pick = function (test) {
    const k = Object.keys(answers).filter(function (t) { return test(t.toLowerCase()); })[0];
    return k ? String(answers[k] == null ? '' : answers[k]).trim() : '';
  };
  const times = pick(function (t) { return t.indexOf('how many') > -1 || t.indexOf('number of times') > -1 || (t.indexOf('times') > -1 && t.indexOf('best time') === -1); });
  const tried = pick(function (t) { return t.indexOf('tried') > -1; });
  const affected = pick(function (t) { return t.indexOf('affected') > -1; });
  const share = pick(function (t) { return t.indexOf('share') > -1 && t.indexOf('impact') > -1; });
  return {
    student_name: pick(function (t) { return t.indexOf('name') > -1 && t.indexOf('800') === -1 && t.indexOf('your') === -1; }),
    student_id: pick(function (t) { return t.indexOf('800') > -1; }),
    behavior_type: pick(function (t) { return t.indexOf('behavior') > -1; }),
    referring_staff: pick(function (t) { return t.indexOf('your name') > -1; }) || respondentEmail || '',
    incident_summary: affected ? 'Who was affected and how: ' + affected : '',
    behavior_pattern: [times && 'Times and dates: ' + times, tried && 'Tried: ' + tried].filter(String).join(' | '),
    source: 'teacher_referral', status: 'suggested',
    referral_json: Object.assign({ teacher_shares_impact: share, best_time_to_reach: pick(function (t) { return t.indexOf('best time') > -1 || t.indexOf('reach') > -1; }) }, answers),
  };
}

function onTrailbackReferral(e) {
  const answers = {};
  e.response.getItemResponses().forEach(function (ir) {
    const v = ir.getResponse();
    answers[ir.getItem().getTitle()] = Array.isArray(v) ? v.join(', ') : v;
  });
  let email = '';
  try { email = e.response.getRespondentEmail(); } catch (err) { /* form doesn't collect it */ }
  try { createCase_(referralToCase_(answers, email), email || 'referral form'); }
  catch (err) { console.error('Trailback referral not saved: ' + err.message + ' ' + JSON.stringify(answers)); throw err; }
}

// ── OTHER TOOLS (server to server) ──────────────────────────────────
const TRAILBACK_API_ACTIONS_ = ['createTrailbackCase', 'trailbackCasesByStudent'];

/** doPost route for the Accountability system and Climb-er: { action, key, payload }. */
function trailbackApi_(body) {
  let key = '';
  try { key = PropertiesService.getScriptProperties().getProperty('TRAILBACK_API_KEY') || ''; } catch (e) { /* none */ }
  if (!key || String(body.key || '') !== key) throw new Error('Wrong or missing Trailback key.');
  const call = { payload: body.payload || {}, _caller: body.caller ? 'tool: ' + String(body.caller).slice(0, 40) : 'another tool' };
  if (body.action === 'createTrailbackCase') {
    const f = Object.assign({}, call.payload, { status: call.payload.status === 'suggested' ? 'suggested' : 'offered' });
    return createCase_(f, call._caller);
  }
  return trailbackCasesByStudent_(call);
}

function showTrailbackKey() {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('TRAILBACK_API_KEY');
  if (!key) { key = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''); props.setProperty('TRAILBACK_API_KEY', key); }
  safeAlert_('Key for the Accountability system and Climb-er (paste it into their Script Properties as TRAILBACK_API_KEY):\n\n' + key +
    '\n\nThey send it to the student (public) /exec link. Keep it out of pages, docs and chat.');
}

// ── TEST CASE (menu) ────────────────────────────────────────────────
function addTrailbackTestCase() {
  const c = createCase_({ case_id: TEST_CASE_PREFIX_ + Utilities.getUuid().slice(0, 6).toUpperCase(), student_name: 'TEST Student (not real)',
    grade: '7', squad: 'Gold', behavior_type: 'Test only', incident_summary: 'Test case: nothing real. Remove it with Remove test cases.',
    source: 'admin_conference', status: 'offered' }, 'test');
  safeAlert_('Test case ' + c.case_id + ' added as "Offered" (no 800#, no Master Log row, so nothing else changes). Remove it with Trail Journal › Trailback › Remove test cases.');
}

function removeTrailbackTestCases() {
  const sh = ensureSchemaTab_(TABS.PANEL_CASES);
  const vals = sh.getDataRange().getValues();
  const idCol = vals[0].indexOf('case_id');
  let n = 0;
  for (let i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][idCol]).indexOf(TEST_CASE_PREFIX_) === 0) { sh.deleteRow(i + 1); n++; }
  }
  safeAlert_(n + ' test case' + (n === 1 ? '' : 's') + ' removed.');
}

function ensurePanelMembersTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(PANEL_MEMBERS_TAB_);
  if (sh) return sh;
  sh = ss.insertSheet(PANEL_MEMBERS_TAB_);
  sh.getRange(1, 1, 1, 4).setValues([['Name', 'Email', 'Role', 'Active']]).setFontWeight('bold').setBackground('#2c4a35').setFontColor('#ffffff');
  sh.getRange(1, 6).setValue('One row per panel member. Role "Student" rows never get the Tuesday brief. Active: FALSE or No skips someone.');
  sh.setFrozenRows(1);
  return sh;
}
