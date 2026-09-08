# Trail Journal — Fix & Feature Plan (for Cowork)

Project: Trail Journal (trailjournal.org) — student restorative reflection app.
Architecture: static HTML on GitHub Pages (repo: matthewschuck-gif/trail-journal) + Google Apps
Script Web App backend + Google Sheet datastore. The frontend calls the backend with
`callBackend({action, table, ...})` POSTed as text/plain to the Apps Script `/exec` URL.
Backend files: 01_Config.gs, 02_AIProxy.gs, 03_Router.gs, 04_FollowupEmails.gs, 05_AdminAuth.gs,
06_PeaksPosters.gs.

## Answers I need from Matt before/while working (fill in)

1. Email recipients for every completed journal: ____________________ (list of addresses)
2. Should that email contain the FULL reflection (every answer verbatim) or the formatted
   office-record narrative it currently sends? (Recommend: full reflection appended below the
   existing narrative.)
3. Where do you currently edit the live code — directly in the GitHub web editor, or a local
   folder? (Needed to find the true current source, since the repo `main` branch is STALE — it
   still holds the old Supabase version while the live site runs the Apps Script version.)
4. Google Sheet URL for the Trail Journal data, and confirm I can open the Apps Script project
   bound to it (Extensions → Apps Script).

## Phase 0 — Recover source truth (BLOCKING — do first)

The deployed site and the repo disagree. Nothing else is safe until fixed.

- [ ] Locate the real current `index.html` (matches the live site: grade picker, location
      picker, staff password gates, Peaks screen, Apps Script APP_URL). Check: local folders,
      GitHub web-editor history, or save the served page from the live site as a last resort.
- [ ] Copy all `.gs` files out of the Apps Script editor into the repo under `apps-script/`
      (01_Config.gs … 06_PeaksPosters.gs) so backend code is version-controlled from now on.
- [ ] Commit current live state to `main` (archive the old Supabase version on a branch named
      `archive/supabase-era` first). Verify GitHub Pages still serves identical content after
      the push.

Acceptance: fresh clone of `main` contains code identical to what trailjournal.org serves.

## Phase 1 — Add "Removal from Lunch" location (small, frontend-only)

Three edits in `index.html`:

- [ ] In the welcome-screen location picker, after the Afterschool Detention button, add
      (match sibling styling exactly):
      `<button class="squad-btn location-btn" data-loc="Removal from Lunch"
        onclick="selectLocation('Removal from Lunch',this)"
        style="background:var(--gold-light);color:var(--gold-dark);border-color:var(--gold)">
        <span class="sq-icon">🚫</span>Removal from Lunch</button>`
- [ ] In the `LOCATION_TIER` map, after `'Afterschool Detention': 'detention',` add
      `'Removal from Lunch': 'detention',`  (same Part 1 + Part 2 journal flow).
- [ ] In `LOCATION_PEAKS_DISPLAY` (and `PEAKS_POSTER_IMAGE` if present), add a
      `'Removal from Lunch'` entry mirroring the `'Lunch Detention'` entry (cafeteria
      expectations are the right Peaks reference).

Acceptance: pick Removal from Lunch on the live site → detention-tier journal runs → the saved
Sheet row and office email both show location "Removal from Lunch".

## Phase 2 — Full-reflection email to a distribution list on every completion

The backend already auto-sends an office-record email on submission (action
`sendOfficeRecordEmail` in the email .gs file, sent via MailApp). Extend it:

- [ ] Add a `office_record_recipients` key to the `site_config` tab (comma-separated addresses,
      from Question 1). Read it in the email function; fall back to the current recipient if
      the key is empty. MailApp accepts a comma-joined `to` string directly.
- [ ] Append a "FULL REFLECTION" section to the email body: every Part 1 / Part 2 / Part 3
      answer verbatim (parse the stored part1/part2/part3 JSON; label each question; plain
      text is fine). Keep the existing narrative/summary blocks above it.
- [ ] (Nice-to-have) Add a field in the admin Content Editor to edit
      `office_record_recipients` without touching the Sheet.

Acceptance: submit a test journal → every listed recipient gets one email containing both the
narrative and every raw answer.

## Phase 3 — Fix the Google Sheet "full reflection" gap

Diagnosis to confirm: the generic `insert` writes only the payload keys that match existing
column headers on the tab — any field without a header is silently dropped. Also check whether
part1_json/part2_json/part3_json are being JSON-stringified (JSON_COLUMNS list in 01_Config.gs).

- [ ] Diff the submit payload keys in `index.html` (the object sent with `action:'insert'`)
      against the reflections tab's header row. Add every missing header via a
      `migrateAddMissingReflectionColumns_()` one-off in 01_Config.gs (use the existing
      `addMissingColumns_` helper), run once from the editor.
- [ ] Confirm the three part JSONs land complete (not truncated, valid JSON in the cell).
- [ ] Add a readable layer: either (a) a `reflection_text` column the backend writes on insert —
      a flattened, human-readable Q&A block — or (b) a second "Readable" tab the insert also
      appends to. Recommend (a): simplest and exportable.

Acceptance: a new test submission shows every answer in the Sheet, and the readable column
contains the entire reflection as plain text.

## Phase 4 — Rebuild the admin dashboard on the Apps Script backend

Root cause of "admin not working": `admin/index.html` is still the Supabase version, reading a
database the form no longer writes to. Migrate it to the same callBackend pattern as the form.

- [ ] Replace the Supabase client calls with `callBackend({action:'query', table:'reflections',
      ...})` (and site_config where the Content Editor reads/writes). Reuse APP_URL/APP_TOKEN
      and the text/plain CORS-avoidance wrapper from `index.html`.
- [ ] Login: swap the old local `doLogin` check for the server-side `checkStaffPassword` action
      (05_AdminAuth.gs already implements per-staff passwords). Keep the same UI.
- [ ] Ensure the router's admin gate (`enforceAdminGate_`) protects reflections reads —
      reflections should be admin-read-only, public-insert.
- [ ] Re-point every dashboard page at Sheet data: Dashboard stats, Submissions list + detail,
      Export to Sheets/CSV (now trivial — the Sheet IS the data; the export page can deep-link
      the Sheet plus keep CSV download), Content Editor, Panel Follow-Ups and Growth Arc (note:
      panel_sessions may still live in Supabase — decide whether to migrate the panel app the
      same way in a later phase, and until then either keep those two pages reading Supabase or
      hide them behind a "coming soon" note).
- [ ] Deployment reminder: ONE Web App deployment only, "Execute as me" / access "Anyone" —
      never a domain-restricted second deployment (it breaks CORS silently).

Acceptance: admin login works with staff password; Submissions shows today's test entries;
Content Editor round-trips a change; no console errors.

## Phase 5 — Close out

- [ ] Push everything to `main`; verify live site.
- [ ] Send Matt a 5-line summary of what changed + the runbook items only he can do in
      Workspace (Script Properties, re-deploy of the Web App after .gs changes — a new
      deployment version is required for backend edits to go live).
- [ ] Note remaining ecosystem items NOT in this plan: BIC URL-param reader, panel app
      migration to Apps Script, CoC widget paste-ins.

## Working rules

- Branch per phase off `main`; `node --check` every edited script block (copy .gs → .js first);
  test each phase on the live site before starting the next.
- Never commit Script Properties values (APP_TOKEN is already public in page source by design;
  API keys and passwords are not).
