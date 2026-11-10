# trail-journal
Student interactive reflection for restorative discipline, plus the Trailback Panel.

## Apps Script pages (live since 2026-11-06)
The pages are Apps Script files in `apps-script/`: Journal, Respond, Admin, Panel and Followup,
plus TrailCss and TjBridge. `08_Web.gs` serves them with `?page=`. trailjournal.org (GitHub Pages,
`main`) only holds the wrappers in `tools/wrappers/`.
- **Backend calls** go through `tjCall` (google.script.run).
- **Two deployments** (every New version goes on both):
  - the student link (Anyone): journal and Your Side
  - the staff link (Anyone within EASD): admin, panel and follow-up. Staff are signed in with
    Google. The per-adult check-in gate passwords are unchanged.
- **800#:** the journal asks for it and saves it as `student_id`. Your Side of the Trail stays
  anonymous, as it promises.
- **Paste pack:** `node tools/build-paste-pack.js` writes `paste-pack.html` with every file.

## Trailback Panel
A restorative alternative to detention: six adults hear a case for 20 minutes on a Wednesday,
agree on a repair plan, and the detention is held. Plan finished: detention cleared. Not
finished: detention served.
- **Cases** (`12_TrailbackCases.gs`, tab `panel_cases`): suggested → offered → consented →
  tj_done → scheduled → heard_approved → repair_in_progress → closed_completed / closed_reverted
  (declined any time before the panel). Admin › **Trailback Cases** moves them one step at a time.
- **The Master Log** on the Accountability Sheet follows: consent → `Held (Trailback)`,
  completed → `Resolved (Trailback)`, reverted or declined after consent → `Scheduled` + a note
  to pick a new date (never booked automatically).
- **The panel** (`Panel.html`, `11_Trailback.gs`) opens on the queue. The AP approves in the
  app (the signed-in account is recorded) before the student comes back in. Check-ins at 3, 10
  and 30 school days (weekends and the **No School Dates** tab skipped). AI helpers get initials only.
- **Clash of Classes:** a completed case asks for a point on Admin › Clash Flags; staff approve
  it. No names or initials ever go to the Clash sheet.
- **Practice:** `trailjournal.org/panel/?preview=true` (linked in Admin) saves nothing.
- **Sheet menu › Trail Journal › Trailback:** timers (Tuesday brief, weekday overdue check),
  email previews (nothing sent), connect the staff referral form, the key for other tools,
  add/remove a test case.
- **Other tools** (Accountability Consequence Entry Form, Climb-er) call `createTrailbackCase`
  and `trailbackCasesByStudent` on the student /exec link with the `TRAILBACK_API_KEY` key.

### Settings keys (Settings tab)
| Key | Default | What it does |
|---|---|---|
| `PANEL_APPROVERS` | blank | Emails allowed to approve a plan. Blank = anyone on the staff link. |
| `TRAILBACK_CLASH_POINTS` | 1 | Points for a completed plan, after approval. |
| `TRAILBACK_EMAILS` | OFF | ON sends the Tuesday brief, overdue check-ins and reverted-case emails. |
| `TRAILBACK_COORDINATOR_EMAIL` | blank | Kelly: overdue check-ins and reverted cases. |
| `TRAILBACK_CONFLICT_DEADLINE` | Wednesday morning, before the panel | Shown in the Tuesday brief. |
| `ACCOUNTABILITY_SHEET_ID` | 1IVFnkQQ… | The Sheet with the Master Log. |
| `TRAILBACK_REFERRAL_FORM_ID` | 19qZF3mn… | The "Trailback Panel Staff Referral" form. |

Tabs added by **Run setup**: `panel_cases`, `clash_requests`, `No School Dates`, `Panel Members`
(Name | Email | Role | Active; Role "Student" never gets the brief).

## Tests
```bash
node tests/backend.test.js
```
```bash
node tests/records.test.js
```
```bash
node tests/trailback.test.js
```
```bash
node tests/cases.test.js
```
```bash
node tests/page-scripts.test.js
```
```bash
node tests/build-preview.js
```
The last one serves a preview at http://localhost:3044/?page=panel (add `&staff=1` to be signed
in, `&preview=true` for the panel's practice mode), with a stand-in backend.
