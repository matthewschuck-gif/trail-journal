/**
 * Trail Journal -- Apps Script backend
 * File 20: the Trailback Plan document, made from what the panel decided.
 *
 * Same layout as the hand-made "Trailback Plan - DR / JL / TG" docs:
 *   header band (Camp Mountaineer | Trailback Panel, My Trailback Plan, name | grade | panel date)
 *   WHAT WE SEE IN YOU   the strengths the panelists named at the close (+ strengths notes)
 *   YOUR CLIMB           the theme sentence, in a box
 *   YOUR PLAN            "This plan takes the place of your …", then a table:
 *                        Staff Signature once Complete | Action Step | Who will help | Details
 *                        (one row per plan action, plus "Show up for your check-ins.")
 *   YOUR PEOPLE AND CHECK-INS   the check-in people, and #1 / #2 / #3 with the week and a date
 *                        line, and what to have ready (the actions due by then)
 *   signature row        Student | Check-in adult | TRAILBACK_PLAN_SIGNER
 *
 * When: by itself when the panel record is submitted (TRAILBACK_PLAN_DOCS ON), or from the menu
 * (Trailback › Make the plan doc for a case…). Making it again makes a new copy; the newest link
 * is saved on the case (plan_doc_url). Docs go in TRAILBACK_PLAN_FOLDER_ID (blank = a "Trailback
 * Plans" folder in the owner's Drive). They are not shared with anyone: share from Drive.
 * Named "Trailback Plan - XX" with the student's initials, like the hand-made ones.
 */

const PLAN_DOC_PURPLE_ = '#490e6f';
const PLAN_DOC_GOLD_ = '#ffe100';

function planDocsOn_() { return String(webSetting_('TRAILBACK_PLAN_DOCS') || 'ON').toUpperCase() !== 'OFF'; }

// "kelly_rigg@easdpa.org" -> the Panel Members tab's name for that email, else "Kelly Rigg".
function staffNameOf_(s) {
  const raw = String(s || '').trim();
  if (!raw) return '';
  const email = emailOf_(raw);
  if (!email) return raw;
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PANEL_MEMBERS_TAB_);
    if (sh && sh.getLastRow() > 1) {
      const hit = sh.getDataRange().getValues().slice(1).filter(function (r) { return String(r[1] || '').trim().toLowerCase() === email; })[0];
      if (hit && String(hit[0] || '').trim()) return String(hit[0]).trim();
    }
  } catch (e) { /* the tab is optional */ }
  return email.split('@')[0].split(/[._-]+/).filter(String).map(function (w) { return w.charAt(0).toUpperCase() + w.slice(1); }).join(' ');
}

function planList_(v) {
  if (typeof v === 'string') { try { v = JSON.parse(v || '[]'); } catch (e) { v = v ? [v] : []; } }
  return Array.isArray(v) ? v : [];
}
function planLongDate_(ymd) {
  const p = String(ymd || '').slice(0, 10).split('-').map(Number);
  if (p.length < 3 || !p[0]) return '';
  const d = new Date(p[0], p[1] - 1, p[2], 12);
  return ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][d.getDay()] + ', ' +
    ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'][p[1] - 1] + ' ' + p[2] + ', ' + p[0];
}
function planWeekOf_(ymd) {
  const p = String(ymd || '').slice(0, 10).split('-').map(Number);
  if (p.length < 3 || !p[0]) return '';
  const d = new Date(p[0], p[1] - 1, p[2], 12);
  d.setDate(d.getDate() - ((d.getDay() + 6) % 7)); // that week's Monday
  return 'Week of ' + ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][d.getMonth()] + ' ' + d.getDate() + ', ' + d.getFullYear();
}
function planSentence_(s) { s = String(s || '').trim(); return s && !/[.!?]$/.test(s) ? s + '.' : s; }

/** Everything the document shows, from the case and its panel session (no Google calls but reads). */
function planDocModel_(c, session) {
  session = session || {};
  const name = String(c.student_name || session.student_name || session.student_initial || '').trim();
  const parts = name.indexOf(',') > -1 ? name.split(',').map(function (x) { return x.trim(); }).reverse() : name.split(/\s+/);
  const display = parts.join(' ').trim();
  const initials = parts.filter(String).map(function (w) { return w.charAt(0).toUpperCase(); }).join('').slice(0, 3) || 'XX';
  const strengths = planList_(session.panel_strengths).map(function (x) { return planSentence_(typeof x === 'string' ? x : x.strength); }).filter(String);
  String(session.strengths_notes || '').split(/\n+/).map(function (x) { return planSentence_(x.replace(/^[-•*]\s*/, '')); }).filter(String)
    .forEach(function (x) { if (strengths.indexOf(x) === -1) strengths.push(x); });
  const actions = repairActionsOf_(c).length ? repairActionsOf_(c) : planList_(session.plan_actions);
  const checkinPeople = [];
  [c.checkin_person || session.followup_person].concat(actions.map(function (a) { return a.adult; })).forEach(function (s) {
    const n = staffNameOf_(s);
    if (n && checkinPeople.indexOf(n) === -1) checkinPeople.push(n);
  });
  const checkinAdult = staffNameOf_(c.checkin_person || session.followup_person) || checkinPeople[0] || '';
  const rows = actions.map(function (a) {
    return { step: planSentence_(a.action), who: staffNameOf_(a.adult), details: planSentence_(a.test || a.details || ''), due: String(a.due || '').slice(0, 10) };
  });
  rows.push({ step: 'Show up for your check-ins.', who: checkinAdult, details: 'Be honest about what is going well and what is hard.', due: '' });
  const dates = { 3: c.checkin_3, 10: c.checkin_10, 30: c.checkin_30 };
  let fd = session.followup_dates || {};
  if (typeof fd === 'string') { try { fd = JSON.parse(fd || '{}'); } catch (e) { fd = {}; } }
  [3, 10, 30].forEach(function (n) { if (!dates[n] && fd['checkin_' + n]) dates[n] = fd['checkin_' + n]; });
  let prev = '';
  const checkins = [3, 10, 30].map(function (n, i) {
    const when = String(dates[n] || '').slice(0, 10);
    const ready = rows.slice(0, -1).filter(function (r) { return r.due && r.due > prev && (!when || r.due <= when); }).map(function (r) { return r.step.replace(/\.$/, '') + ' done.'; });
    if (!ready.length) ready.push(i === 0 ? 'Your first steps on the calendar.' : i === 1 ? 'Show what you have finished so far.' : 'How school is going, and what you are proud of.');
    prev = when || prev;
    return { label: '#' + (i + 1) + '  (' + (n === 3 ? 'next school day' : n + ' school days') + ')', when: when ? planWeekOf_(when) : '', ready: ready };
  });
  const consequence = String(c.consequence_type || '').replace(/^ADMIN ONLY\s*/i, '').replace(/\s*\((LD|ASD)\)/i, '').replace(/\s+(M-F|TR)$/i, '').trim();
  const takes = String(webSetting_('TRAILBACK_PLAN_TAKES_PLACE') || 'This plan takes the place of your {consequence}. Following through is what keeps it that way.')
    .replace('{consequence}', consequence ? consequence.toLowerCase().replace(/\b(ld|asd|iss|oss)\b/g, function (m) { return m.toUpperCase(); }) : 'consequence');
  return {
    title: 'Trailback Plan - ' + initials, student: display, grade: String(c.grade || session.grade || '').replace(/\D/g, ''),
    panelDate: planLongDate_(c.scheduled_date || session.panel_date), strengths: strengths,
    climb: planSentence_(session.theme_sentence || ''), takes: takes, rows: rows, checkinPeople: checkinPeople, checkinAdult: checkinAdult,
    checkins: checkins, signer: String(webSetting_('TRAILBACK_PLAN_SIGNER') || 'Mr. Schuck'),
  };
}

function planDocFolder_() {
  const id = String(webSetting_('TRAILBACK_PLAN_FOLDER_ID') || '').match(/[-\w]{20,}/);
  if (id) return DriveApp.getFolderById(id[0]);
  const it = DriveApp.getFoldersByName('Trailback Plans');
  return it.hasNext() ? it.next() : DriveApp.createFolder('Trailback Plans');
}

function styleCell_(cell, bg, color, bold, size) {
  if (bg) cell.setBackgroundColor(bg);
  const t = cell.editAsText();
  if (color) t.setForegroundColor(color);
  if (bold !== undefined) t.setBold(bold);
  if (size) t.setFontSize(size);
  return cell;
}

/** Writes the Google Doc. Returns { url, id, title }. */
function writePlanDoc_(m) {
  const doc = DocumentApp.create(m.title), body = doc.getBody();
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(48).setMarginRight(48);
  const H = function (text) {
    const p = body.appendParagraph(text);
    p.editAsText().setBold(true).setFontSize(13).setForegroundColor(PLAN_DOC_PURPLE_);
    p.setSpacingBefore(14).setSpacingAfter(4);
    return p;
  };
  // Header band.
  const head = body.appendTable([['CAMP MOUNTAINEER  |  TRAILBACK PANEL\nMy Trailback Plan\n' +
    [m.student, m.grade ? 'Grade ' + m.grade : '', m.panelDate ? 'Panel: ' + m.panelDate : ''].filter(String).join('  |  ')], ['Ephrata Middle School  |  Together We Climb']]);
  styleCell_(head.getCell(0, 0), PLAN_DOC_PURPLE_, '#ffffff', true, 12);
  head.getCell(0, 0).editAsText().setFontSize(0, 'CAMP MOUNTAINEER  |  TRAILBACK PANEL'.length, 9);
  styleCell_(head.getCell(1, 0), PLAN_DOC_GOLD_, PLAN_DOC_PURPLE_, true, 10);
  if (body.getChild(0).getType() === DocumentApp.ElementType.PARAGRAPH && !body.getChild(0).asParagraph().getText()) body.removeChild(body.getChild(0));

  H('WHAT WE SEE IN YOU');
  (m.strengths.length ? m.strengths : ['(The panel adds the strengths it named here.)']).forEach(function (s) { body.appendListItem(s).setGlyphType(DocumentApp.GlyphType.BULLET); });

  H('YOUR CLIMB');
  const climb = body.appendTable([[m.climb || '(The panel\'s theme for this climb.)']]);
  styleCell_(climb.getCell(0, 0), '#f3edf8', PLAN_DOC_PURPLE_, true, 12);

  H('YOUR PLAN');
  body.appendParagraph(m.takes).editAsText().setItalic(true).setFontSize(10);
  const plan = body.appendTable([[' Staff Signature once Complete', ' Action Step', 'Who will help you complete this action step?', ' Details']]
    .concat(m.rows.map(function (r) { return [' ', r.step, r.who || ' ', r.details || ' ']; })));
  for (let j = 0; j < 4; j++) styleCell_(plan.getCell(0, j), PLAN_DOC_PURPLE_, '#ffffff', true, 10);
  for (let i = 1; i < plan.getNumRows(); i++) {
    styleCell_(plan.getCell(i, 1), null, '#231018', true, 10);
    styleCell_(plan.getCell(i, 2), null, '#231018', false, 10);
    styleCell_(plan.getCell(i, 3), null, '#231018', false, 10);
  }

  H('YOUR PEOPLE AND CHECK-INS');
  const people = body.appendParagraph('Your check-in people: ' + (m.checkinPeople.join(' and ') || '________________'));
  people.editAsText().setBold(0, 'Your check-in people:'.length - 1, true);
  const ck = body.appendTable([['CHECK-IN', 'WHEN', 'WHAT TO HAVE READY']].concat(m.checkins.map(function (x) {
    return [x.label, (x.when ? x.when + '\n' : '') + 'Date: ____________', x.ready.map(function (r) { return '• ' + r; }).join('\n')];
  })));
  for (let j = 0; j < 3; j++) styleCell_(ck.getCell(0, j), PLAN_DOC_PURPLE_, '#ffffff', true, 10);
  for (let i = 1; i < ck.getNumRows(); i++) { styleCell_(ck.getCell(i, 0), null, PLAN_DOC_PURPLE_, true, 10); styleCell_(ck.getCell(i, 1), null, null, false, 10); styleCell_(ck.getCell(i, 2), null, null, false, 10); }

  body.appendParagraph(' ');
  const sig = body.appendTable([['\n\n Student', '\n\n Check-in adult', '\n\n ' + m.signer]]);
  for (let j = 0; j < 3; j++) styleCell_(sig.getCell(0, j), null, '#231018', false, 10);
  doc.saveAndClose();
  const file = DriveApp.getFileById(doc.getId());
  try { file.moveTo(planDocFolder_()); } catch (e) { console.warn('Plan doc stays in My Drive: ' + e.message); }
  return { url: doc.getUrl(), id: doc.getId(), title: m.title };
}

/** Makes the doc for one case and saves the link on it. */
function makePlanDoc_(caseId, by) {
  const c = getCase_(caseId);
  if (!c) throw new Error('No Trailback case ' + caseId + '.');
  const session = c.panel_session_id ? panelRow_(c.panel_session_id) : null;
  if (!repairActionsOf_(c).length && !(session && planList_(session.plan_actions).length)) throw new Error('Case ' + caseId + ' has no plan yet: the panel locks the plan first.');
  const out = writePlanDoc_(planDocModel_(c, session));
  saveCase_(c, { plan_doc_url: out.url }, by || 'plan doc', c.status, 'Plan doc made: ' + out.title);
  return out;
}

// From casePanelSubmitted_ (12_TrailbackCases.gs). Never stops the submit.
function autoPlanDoc_(caseId, by) {
  if (!caseId || !planDocsOn_()) return;
  try { makePlanDoc_(caseId, by); } catch (e) { console.warn('Plan doc not made for ' + caseId + ': ' + e.message); }
}

// Staff link action (03_Router.gs): { case_id } -> { url, title }.
function makePlanDocAction_(body) { return makePlanDoc_(String((body.payload || {}).case_id || ''), whoIs_(body)); }

// Menu: Trailback › Make the plan doc for a case…
function menuMakePlanDoc() {
  const ui = SpreadsheetApp.getUi();
  const open = allCases_().filter(function (c) { return c.status === 'repair_in_progress' || c.status === 'heard_approved'; });
  const r = ui.prompt('Make a Trailback Plan doc', 'Type the case ID (from the panel_cases tab, e.g. TB-…), or ALL for every open plan without a doc.\n\nOpen plans: ' +
    (open.map(function (c) { return c.case_id + (c.plan_doc_url ? ' (has one)' : ''); }).join(', ') || 'none'), ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const want = r.getResponseText().trim();
  const ids = /^all$/i.test(want) ? open.filter(function (c) { return !c.plan_doc_url; }).map(function (c) { return c.case_id; }) : [want];
  const done = [], failed = [];
  ids.forEach(function (id) { try { done.push(id + ': ' + makePlanDoc_(id, Session.getActiveUser().getEmail() || 'menu').url); } catch (e) { failed.push(id + ': ' + e.message); } });
  ui.alert('Trailback Plan docs', (done.length ? 'Made:\n' + done.join('\n') : 'None made.') + (failed.length ? '\n\nNot made:\n' + failed.join('\n') : ''), ui.ButtonSet.OK);
}
