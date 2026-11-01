# trail-journal
Student interactive reflection for restorative discipline 

## Apps Script pages (built 2026-11-01, not live yet)
The five pages also exist as Apps Script copies in `apps-script/`: Journal, Respond, Admin,
Panel and Followup, plus TrailCss and TjBridge. `08_Web.gs` serves them with `?page=`.
- **Backend calls** go through `tjCall` (google.script.run) instead of the public token.
- **Two deployments:**
  - the student link (Anyone): journal and Your Side
  - the staff link (Anyone within EASD): admin, panel and follow-up. Staff are signed in with
    Google, so there's no shared password. The per-adult check-in gate passwords are unchanged.
- **800#:** the journal asks for it and saves it as `student_id`. Your Side of the Trail stays
  anonymous, as it promises.
- **The switch:** `tools/wrappers/` becomes the site's pages:
  - the root and `respond/` show the student link in a frame
  - `admin/`, `panel/` and `panel/followup/` go to the staff link
- **After the switch:** set Settings › `LEGACY_API` = OFF.
- **The GitHub pages** (index.html, admin/ …) stay as they are until then.
  `tools/split-for-apps-script.js` made the copies; it won't overwrite them without `--force`.

Tests:
```bash
node tests/backend.test.js
```
```bash
node tests/build-preview.js
```
The second serves a preview at http://localhost:3044/?page=journal, with a stand-in backend.
