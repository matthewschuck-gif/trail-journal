/**
 * Trail Journal -- Apps Script backend
 * File 21: the Staff Access list, the Staff Actions log, security alerts and nightly backups.
 *
 * Two kinds of staff on the staff link ("Anyone within EASD"):
 *   - Any school account at STAFF_DOMAIN counts as staff for the everyday staff pages: panel
 *     sign-up, check-ins, Give / Get. (body._staffEmail, as before.)
 *   - The admin tools (admin dashboard, Trailback Panel, follow-up scheduler, every staff-only
 *     action in 05_AdminAuth.gs) need the person on the Staff Access tab with Active ticked
 *     (Role: Admin, Counselor, Secretary, Teacher). The Sheet owner always counts. (body._staffAdmin)
 *   Settings STAFF_ACCESS = LIST (default) turns that on; DOMAIN = the old way (any school account).
 *
 * The shared ADMIN_PASSWORD still opens the admin tools while PASSWORD_SIGN_IN is ON (default),
 * so nobody is locked out before the list is filled in. Turn it OFF afterwards. The per-adult
 * check-in gate passwords inside the student journal are a separate thing and unchanged.
 *
 * Setup seeds Staff Access with the owner and the emails already trusted in Settings
 * (PANEL_ADMINS, PANEL_APPROVERS, STAFF_EMAILS, TRAILBACK_COORDINATOR_EMAIL), all Active.
 *
 * STAFF ACTIONS tab: password sign-ins, refused passwords, sign-outs, and every staff-only change
 * (reads are not logged). ALERTS (SECURITY_ALERT_EMAILS, blank = the owner; once an hour each):
 * someone's first time in the admin tools, 5 and 10 wrong staff passwords.
 * BACKUP: Trail Journal menu › Turn on nightly backup ("Trail Journal Backups", BACKUP_KEEP copies).
 */

const STAFF_ACCESS_TAB_ = 'Staff Access';
const STAFF_ACCESS_HEADERS_ = ['Email', 'Name', 'Role', 'Active', 'Last Sign-in'];
const STAFF_ACTIONS_TAB_ = 'Staff Actions';
const STAFF_ACTIONS_HEADERS_ = ['Time', 'Who', 'Email', 'Signed in with', 'Action', 'Detail'];
const STAFF_ROLES_ = ['Admin', 'Counselor', 'Secretary', 'Teacher'];
const ACCESS_DEFAULTS_ = [
  ['STAFF_ACCESS', 'LIST', 'LIST = the admin tools (admin dashboard, Trailback Panel, follow-up scheduler) need the person on the Staff Access tab with Active ticked. DOMAIN = any school account (the old way). Sign-up, check-ins and Give / Get stay open to every school account.'],
  ['PASSWORD_SIGN_IN', 'ON', 'ON = the shared staff password (ADMIN_PASSWORD) still opens the admin tools. Turn OFF once everyone who needs them is on Staff Access.'],
  ['SECURITY_ALERT_EMAILS', '', 'Who gets security alert emails (comma separated). Blank = the Sheet owner.'],
  ['BACKUP_KEEP', '14', 'Nightly backup copies of this Sheet to keep in the Trail Journal Backups folder.'],
];
// Staff-only actions that only read: not written to Staff Actions.
const READ_ONLY_STAFF_ACTIONS_ = ['followupDates', 'clashRequests', 'trailbackCases', 'trailbackCase', 'trailbackCasesByStudent'];

function accessSetting_(key) {
  const v = webSetting_(key);
  if (v !== '') return v;
  const d = ACCESS_DEFAULTS_.filter(function (x) { return x[0] === key; })[0];
  return d ? d[1] : '';
}
function ownerEmail_() { try { return String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase(); } catch (e) { return ''; } }
function passwordSignInOn_() { return String(accessSetting_('PASSWORD_SIGN_IN')).toUpperCase() !== 'OFF'; }

// { email, name, role, row } for someone on Staff Access with Active ticked (or the owner), else null.
function staffAccessRecord_(email) {
  email = String(email || '').trim().toLowerCase();
  if (!email) return null;
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STAFF_ACCESS_TAB_);
  let rec = null;
  if (sh && sh.getLastRow() > 1) {
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, STAFF_ACCESS_HEADERS_.length).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toLowerCase() === email) { rec = { row: i + 2, name: String(rows[i][1] || '').trim(), role: String(rows[i][2] || '').trim(), active: rows[i][3] === true || /^(true|yes|y|x)$/i.test(String(rows[i][3]).trim()) }; break; }
    }
  }
  if (email === ownerEmail_()) return { email: email, name: (rec && rec.name) || 'Sheet owner', role: (rec && rec.role) || 'Admin', row: rec ? rec.row : 0 };
  if (!rec || !rec.active) return null;
  return { email: email, name: rec.name || email, role: rec.role || 'Staff', row: rec.row };
}

// The signed-in school account if it may use the admin tools, else ''.
function staffAdminEmail_() {
  const email = staffEmail_(); // 08_Web.gs: a school account at STAFF_DOMAIN (or '')
  if (!email) return '';
  if (String(accessSetting_('STAFF_ACCESS')).toUpperCase() === 'DOMAIN') return email;
  // Until Run setup has made the Staff Access tab, nothing changes (any school account).
  if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STAFF_ACCESS_TAB_)) return email;
  return staffAccessRecord_(email) ? email : '';
}

// Page load on the staff link: first-time alert and Last Sign-in for people on the list.
function noteStaffVisit_(email, page) {
  if (!email || ['admin', 'panel', 'followup', 'outcomes'].indexOf(page) === -1) return;
  try {
    const rec = staffAccessRecord_(email);
    if (!rec) return;
    const cache = CacheService.getScriptCache(), k = 'tj-visit-' + email;
    if (cache.get(k)) return; // once an hour is plenty
    cache.put(k, '1', 3600);
    const props = PropertiesService.getScriptProperties(), seen = 'tj-seen-' + email;
    if (!props.getProperty(seen)) {
      props.setProperty(seen, nowIso_());
      if (email !== ownerEmail_()) securityAlert_('first-' + email, 'First sign-in: ' + rec.name, [rec.name + ' (' + email + ') opened the Trail Journal ' + page + ' page with their school account for the first time.', 'Their Staff Access role: ' + rec.role + '. Untick Active if they should not have it.']);
    }
    if (rec.row) SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STAFF_ACCESS_TAB_).getRange(rec.row, 5).setValue(Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd HH:mm'));
    logStaffAction_(rec.name, email, 'Google account', 'Opened ' + page, '');
  } catch (e) { console.warn('Staff visit: ' + e); }
}

function logStaffAction_(who, email, via, action, detail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(STAFF_ACTIONS_TAB_);
    if (!sh) {
      sh = ss.insertSheet(STAFF_ACTIONS_TAB_);
      sh.getRange(1, 1, 1, STAFF_ACTIONS_HEADERS_.length).setValues([STAFF_ACTIONS_HEADERS_]).setFontWeight('bold').setBackground('#2c4a35').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
    sh.appendRow([Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd HH:mm'), who || '', email || '', via || '', String(action || '').slice(0, 200), String(detail || '').slice(0, 300)]);
  } catch (e) { console.warn('Staff Actions log: ' + e); }
}

// After the admin gate lets a request through: log what changed (never the record itself).
function logGatedAction_(body) {
  if (body.action === 'query' || READ_ONLY_STAFF_ACTIONS_.indexOf(body.action) > -1) return;
  const rec = body._staffAdmin ? staffAccessRecord_(body._staffAdmin) : null;
  logStaffAction_(rec ? rec.name : (body._staffEmail ? body._staffEmail : 'Shared password'), body._staffAdmin || body._staffEmail || '',
    body._staffAdmin || body._staffEmail ? 'Google account' : 'Staff password', String(body.action || '') + (body.table ? ' ' + body.table : ''),
    body.payload && body.payload.case_id ? 'case ' + body.payload.case_id : body.match && body.match.id ? 'row ' + body.match.id : '');
}

function securityAlert_(key, subject, lines) {
  try {
    const cache = CacheService.getScriptCache(), k = 'tj-alert-' + key;
    if (cache.get(k)) return false;
    cache.put(k, '1', 3600);
    const to = String(accessSetting_('SECURITY_ALERT_EMAILS')).split(',').map(function (s) { return s.trim(); }).filter(Boolean);
    if (!to.length && ownerEmail_()) to.push(ownerEmail_());
    if (!to.length) return false;
    const e = function (s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
    MailApp.sendEmail({ to: to.join(','), subject: 'Trail Journal security: ' + subject, name: 'Trail Journal',
      htmlBody: '<div style="font-family:Arial,sans-serif;max-width:560px;border-top:6px solid #490e6f;padding:16px">' +
        '<div style="font-size:12px;color:#490e6f;font-weight:bold;letter-spacing:1px">TRAIL JOURNAL · SECURITY</div>' +
        '<h2 style="margin:6px 0 12px;color:#222">' + e(subject) + '</h2>' +
        lines.map(function (l) { return '<p style="margin:6px 0;color:#333">' + e(l) + '</p>'; }).join('') +
        '<p style="margin-top:14px;color:#777;font-size:12px">Full record: the Staff Actions tab. Together We Climb.</p></div>' });
    return true;
  } catch (err) { console.warn('Security alert: ' + err); return false; }
}

// Setup: Settings rows, the Staff Access tab seeded with the trusted emails, Staff Actions.
function setupStaffAccess_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(WEB_SETTINGS_TAB_);
  if (settings) {
    const have = {};
    settings.getDataRange().getValues().slice(1).forEach(function (r) { have[String(r[0]).trim()] = true; });
    const missing = ACCESS_DEFAULTS_.filter(function (d) { return !have[d[0]]; });
    if (missing.length) settings.getRange(settings.getLastRow() + 1, 1, missing.length, 3).setValues(missing);
  }
  let sh = ss.getSheetByName(STAFF_ACCESS_TAB_);
  if (!sh) {
    sh = ss.insertSheet(STAFF_ACCESS_TAB_);
    sh.getRange(1, 1, 1, STAFF_ACCESS_HEADERS_.length).setValues([STAFF_ACCESS_HEADERS_]).setFontWeight('bold').setBackground('#2c4a35').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    const seen = {}, rows = [];
    [ownerEmail_()].concat(['PANEL_ADMINS', 'PANEL_APPROVERS', 'STAFF_EMAILS', 'TRAILBACK_COORDINATOR_EMAIL'].reduce(function (all, k) {
      return all.concat(String(webSetting_(k)).split(',').map(function (s) { return s.trim().toLowerCase(); }));
    }, [])).forEach(function (email) {
      if (!/@/.test(email) || seen[email]) return;
      seen[email] = true;
      rows.push([email, '', 'Admin', true, '']);
    });
    if (rows.length) sh.getRange(2, 1, rows.length, STAFF_ACCESS_HEADERS_.length).setValues(rows);
  }
  try {
    const n = Math.max(sh.getMaxRows() - 1, 1);
    sh.getRange(2, 3, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STAFF_ROLES_, true).build());
    sh.getRange(2, 4, n, 1).insertCheckboxes();
  } catch (e) { /* fine in tests */ }
  if (!ss.getSheetByName(STAFF_ACTIONS_TAB_)) logStaffAction_('', '', '', 'Staff Actions log started', '');
}

// ── NIGHTLY BACKUP ─────────────────────────────────────────────────────
function backupTrailJournal() {
  const props = PropertiesService.getScriptProperties();
  let folder = null;
  const fid = props.getProperty('tj-backup-folder');
  if (fid) { try { folder = DriveApp.getFolderById(fid); if (folder.isTrashed && folder.isTrashed()) folder = null; } catch (e) { folder = null; } }
  if (!folder) { folder = DriveApp.createFolder('Trail Journal Backups'); props.setProperty('tj-backup-folder', folder.getId()); }
  DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).makeCopy('Trail Journal backup ' + Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd'), folder);
  const keep = Math.max(1, Number(accessSetting_('BACKUP_KEEP')) || 14), files = [], it = folder.getFiles();
  while (it.hasNext()) { const f = it.next(); if (/^Trail Journal backup /.test(f.getName())) files.push(f); }
  files.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  files.slice(keep).forEach(function (f) { f.setTrashed(true); });
  return folder.getUrl();
}

function menuTurnOnBackup() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'backupTrailJournal') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('backupTrailJournal').timeBased().everyDays(1).atHour(2).create();
  safeAlert_('Nightly backup is on (about 2 a.m.). The first copy is in the Trail Journal Backups folder:\n' + backupTrailJournal());
}
