// One-off: every journal is saved to the Sheet from its first screen (index.html + Journal.html).
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');

const oldSave = `async function saveRow(row) {
  try {
    if (S._partialSaved) {
      await callBackend({ action: 'update', table: 'reflections', match: { session_id: row.session_id }, patch: row });
    } else {
      await callBackend({ action: 'insert', table: 'reflections', row: row });
    }
  } catch(e) { console.error('Save error:', e); }
}`;

const newSave = `async function saveRow(row) {
  try {
    // One server action finds this session's row (from a progress save, Ran Out of Time or an
    // earlier try) or makes it -- no insert/update guessing, no duplicate rows (10_Records.gs).
    await callBackend({ action: 'upsertReflection', row: row });
    S._partialSaved = true;
  } catch(e) {
    if (!/Unknown action/.test(String(e && e.message))) { console.error('Save error:', e); return; }
    // Backend not updated yet: the old way.
    try {
      if (S._partialSaved) {
        await callBackend({ action: 'update', table: 'reflections', match: { session_id: row.session_id }, patch: row });
      } else {
        await callBackend({ action: 'insert', table: 'reflections', row: row });
      }
      S._partialSaved = true;
    } catch(e2) { console.error('Save error:', e2); }
  }
}

// ─── PROGRESS SAVES ─────────────────────────────────────────────────────────
// Every screen change after the welcome screen saves what's written so far to the Sheet as
// status "in_progress", so a journal is on the Sheet even if the device dies, the tab is
// closed without the exit save getting out, or nobody clicks Ran Out of Time. The server's
// safety net emails an in-progress journal nobody has touched for an hour as INCOMPLETE.
// Saves run one at a time; a change while one is out waits for it.
let progressSaving_ = null, progressAgain_ = false;
function saveProgress_(n) {
  if (IS_PREVIEW || S._submitted || S._exitBeaconSent || !S.name || !S.sessionId || n === 0 || n === 'done') return;
  if (progressSaving_) { progressAgain_ = true; return; }
  const p1 = getPart1(), p2 = getPart2(), p3 = getPart3();
  const row = {
    session_id: S.sessionId, initial: S.initial,
    student_name: S.name, grade: S.grade, squad: S.squad || null, team: S.team || null,
    location: S.location, peak: S.peak,
    cm_reason: S.cmReason, cm_initiated_by: S.cmInitiatedBy, cm_why: S.cmWhy,
    part1_json: p1, part2_json: p2, part3_json: p3,
    prior_count: S.priorCount || 0, status: 'in_progress',
    gate0_unlocked_by: S.gate0StaffName || null,
    gate1_unlocked_by: S.gate1StaffName || null, gate2_unlocked_by: S.gate2StaffName || null,
  };
  if (S.studentId) row.student_id = S.studentId;
  progressSaving_ = callBackend({ action: 'upsertReflection', row: row })
    .then(() => { S._partialSaved = true; })
    .catch(e => console.error('Progress save error:', e))
    .then(() => { progressSaving_ = null; if (progressAgain_) { progressAgain_ = false; saveProgress_(n); } });
}`;

['index.html', 'apps-script/Journal.html'].forEach(f => {
  const file = path.join(root, f);
  let t = fs.readFileSync(file, 'utf8');
  const crlf = t.includes('\r\n'); if (crlf) t = t.replace(/\r\n/g, '\n');
  if (t.includes('function saveProgress_(')) { console.log(f + ': already patched'); return; }
  const swap = (a, b) => { if (!t.includes(a)) throw new Error(f + ': missing ' + a.slice(0, 50)); t = t.replace(a, b); };
  swap(oldSave, newSave);
  swap("  document.getElementById('rot-btn').classList.toggle('show', !noMeter.includes(n));\n  saveDraft();\n}",
       "  document.getElementById('rot-btn').classList.toggle('show', !noMeter.includes(n));\n  saveDraft();\n  saveProgress_(n);\n}");
  swap("  if (S.tier === 'full' && !validatePart3()) return;\n  clearDraft();\n  goTo('done');",
       "  if (S.tier === 'full' && !validatePart3()) return;\n  S._submitted = true; // no more progress saves; the submit below is the final one\n  clearDraft();\n  goTo('done');");
  fs.writeFileSync(file, crlf ? t.replace(/\n/g, '\r\n') : t);
  console.log(f + ': patched');
});
