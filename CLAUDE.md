# Trail Journal — handoff from a Cowork (cloud) session

This file is here so a local Claude Code session in this folder has full context
without Matt having to re-explain everything. It documents where things actually
stand as of this handoff — read it before doing anything else in this repo.

## What this project is

Trail Journal (trailjournal.org) is a student restorative-reflection app at
Ephrata Middle School. Architecture: static HTML on GitHub Pages (this repo,
`matthewschuck-gif/trail-journal`) + a Google Apps Script Web App backend + a
Google Sheet datastore. The frontend calls the backend with
`callBackend({action, table, ...})` POSTed as `text/plain` to the Apps Script
`/exec` URL (works around a CORS preflight limitation). Backend files live in
`apps-script/`: `01_Config.gs`, `02_AIProxy.gs`, `03_Router.gs`,
`04_FollowupEmails.gs`, `05_AdminAuth.gs`, `06_PeaksPosters.gs`.

The original task doc Matt uploaded is at `docs/trailjournal-cowork-plan.md` in
this repo — read that too, it has the full 5-phase plan. This file only covers
what's changed *since* that doc was written and corrects a couple of things it
got wrong.

## Branch truth — READ THIS FIRST

The repo has three branches and they are **not** what a fresh clone's default
branch suggests:

- **`main`** — STALE. Still the old Supabase-backed version of the app. Do not
  branch off this or assume it reflects the live site.
- **`archive/supabase-original`** — a frozen rollback copy of that same old
  Supabase state. Leave it alone.
- **`migration/google-appscript`** — this is the real, current source. All 5
  `.gs` backend files plus all 5 rewired frontend HTML files live here, and per
  Matt, GitHub Pages is currently serving *this branch directly* (not `main`).
  **Branch off this one, not `main`, for any new work**, unless/until the
  Phase 0 cutover below happens. (A handful of leftover comments in
  `index.html` still say "Supabase" — e.g. near the goal/insight-level save
  call — these are stale comment labels from before the migration, not actual
  Supabase calls; everything routes through `callBackend` now. Harmless, but
  worth cleaning up if you're in that area.)

There's also `MIGRATION_RUNBOOK.md` on `migration/google-appscript` (written by
whichever session did the Supabase→Apps Script port) documenting the one-time
Google Sheet + Apps Script setup Matt already did to get the site live on this
backend. Worth reading for how the Apps Script side is provisioned.

### Outstanding Phase 0 decision (not yet made)

The original plan's Phase 0 wanted `main` to match what's actually live. That's
now just: merge `migration/google-appscript` into `main` (no code changes
needed — it's a fast-forward/no-conflict merge), then decide whether to
repoint GitHub Pages from the branch to `main`. **Ask Matt which he wants**
before doing this — he hadn't decided as of this handoff. It's a two-minute
job either via `gh pr create`/`gh pr merge` or the GitHub web UI; no need to
overthink it once he answers.

## What's already done (this session, on `migration/google-appscript`)

Two commits ahead of `origin/migration/google-appscript` (bundled alongside
this file — see below for how to pull it in).

### Commit 1: "Add Removal from Lunch location (full-tier reflection)"

- New welcome-screen location button next to the detention buttons
  (`index.html`), mapped to tier `'full'` in `LOCATION_TIER_` —
  **deliberately not `'detention'`**. Matt wants the complete Part 1 + Part 2
  + Part 3 reflection for this location.
  - Correction to the original plan doc: it also said to add entries to
    `LOCATION_PEAKS_DISPLAY_` / `PEAKS_POSTER_IMAGE_` for the new location.
    Don't — those two maps are keyed off Part 1's "Where did it happen?"
    incident-location dropdown (Cafeteria already exists there), which is a
    *different* concept from this welcome-screen consequence-location picker.
- Added "Removal from Lunch" to the admin manual-entry Location `<select>`
  (`admin/index.html`). Its tier auto-fill JS only special-cases
  Office/ISS/Camp Mountaineer, so it correctly falls through to the
  already-selected `full` default — no JS change needed there.

**Why this is a patch/bundle instead of already pushed**: the Cowork cloud
sandbox this session ran in blocks git pushes to repos not in that session's
pre-authorized list, regardless of credentials supplied. That's a sandbox
policy, not a git or auth problem — Claude Code running locally (like you,
presumably, if you're reading this) doesn't have that restriction, since it's
using the real local git config and network.

### Removal from Lunch ↔ "Lunchtime Redirection Reflection" Google Form — verified mapping

Matt shared a Google Form (title: "Lunchtime Redirection Reflection", 21
questions) as the content reference for this location. This was checked
line-by-line against the existing `full` tier flow, not just skimmed:

| Form question | Already in the app? |
|---|---|
| First/Last Name, Grade | Yes — collected earlier in the flow (`student_name`, `grade`) |
| Today's Date | Yes — `created_at` timestamp, not a form field |
| Which Expectation (Be Present/Personable/Productive) | Yes — exact match, the Peaks picker (`selectPeak`, `S.peak`) |
| What Happened / Who was involved / Witnessed / Where / What were you thinking / When | Yes — `q-who`, `q-witness`, `q-where`, `q-think`, `q-when` on screen-1 (Part 1), near-identical wording already |
| "Where did it happen" options (Hallway/Classroom/Bus/Cafeteria/Other) | Yes, and more thorough — `q-where` already has Classroom, Hallways, Cafeteria, Auditorium, Bathroom, Media Center, "Traveling To & From School" (covers Bus), Digital Environment, Other |
| Why did you do it / reasoning | Close — covered by the existing goal/reasoning prompt ("What was your goal? What were you trying to get, prove, or avoid?"), not identical wording |
| Emotions at the time / now, change Yes/No + why | Yes — the before/during and now emotion-zone pickers plus a comparison screen already do this |
| Write 3 solutions / pick best one / follow-through plan / what you'll do differently | **Different mechanic, not a literal match.** The app's Part 3 ("Making Things Right") uses structured restorative-justice checkboxes across categories (clean a space, act of kindness, apology, skill-building, creative project) plus a separate "Growth Plan & Follow-Through" section (when you'll finish, what you need, who you'll show it to, what you'll do if you get stuck) — richer than three free-text ideas, but not the same shape as the form. Left as-is in this patch; confirm with Matt before changing it — this is a deliberate existing design, not an oversight. |
| **Team** (M/O/U/N/T/S/Building MTSS Team) | **Not collected anywhere in Trail Journal.** Not the Clash of Classes squad field (`S.squad`: BLACK/GOLD/GREY/PURPLE — a different system) or anything else in the app. Genuinely new if it's actually needed. Ask Matt whether "Removal from Lunch" needs to capture this, or whether it's specific to how the Building MTSS Team uses their standalone form for their own purposes. |

Net: the tier=`'full'` change already gets ~90% of the form's content for
free via existing fields. The two real open items (Team field, whether Part
3 should more literally match the form's "3 solutions" shape) need Matt's
answer, not an assumption — don't silently build either one.

## Phase 2 — full-reflection email to a distribution list (spec verified against actual code)

The original plan doc's Phase 2 assumed the backend needs to parse stored
`part1_json`/`part2_json`/`part3_json` strings to build the full-reflection
email section. **That's not necessary** — checked the actual call site:

- `index.html` around the submit flow calls
  `callBackend({ action: 'sendOfficeRecordEmail', payload: { studentLabel, grade, location, tier, summaryText: buildEmailBody(result||{}, p1, p2, p3) } })`.
  `p1`/`p2`/`p3` are already live JS objects at that point (pre-stringification)
  — `buildEmailBody()` (client-side function in `index.html`) is what currently
  turns them into the narrative summary. Extending *that* function to also
  append a verbatim, labeled Q&A section is simpler than parsing JSON
  server-side, and matches how the existing summary is already built.
- Server side, `sendOfficeRecordEmail_` in `apps-script/04_FollowupEmails.gs`
  (routed from `03_Router.gs`'s `case 'sendOfficeRecordEmail'`) just takes
  whatever `summaryText` it's given and renders it via `formatSummaryHtml_` —
  no change needed there for the content itself.
- The recipient list is the one real backend change: `getOfficeNotifyEmail_()`
  (same file) currently returns a single address from Script Property
  `OFFICE_NOTIFY_EMAIL`, falling back to `emsoffice@easdpa.org`. Needs to
  become a comma-joined list sourced from `site_config` (the existing generic
  key/value Sheet tab already used elsewhere, e.g.
  `peaks_poster_override_<location>` in `06_PeaksPosters.gs`) under a new key
  like `office_record_recipients`. **Note**: `01_Config.gs` only has
  `setSiteConfigValue_`/`deleteSiteConfigValue_` — there's no read helper yet,
  so add a `getSiteConfigValue_(key, fallback)` alongside them (same
  data-range-scan pattern) rather than reinventing the lookup elsewhere.
  `MailApp.sendEmail({ to: '...' })` already accepts a comma-joined string
  directly, matching what the original plan assumed.
- This is the same MailApp + `emailShell_`/`emailInfoRow_` HTML-component
  pattern already used throughout `04_FollowupEmails.gs` — Matt referenced
  wanting this consistent with "other appscript projects"; this file already
  *is* that pattern, so Phase 2 is an extension of it, not a new approach.
- (Nice-to-have from the original plan) an admin Content Editor field to edit
  `office_record_recipients` without touching the Sheet — straightforward
  once `getSiteConfigValue_` exists, using the same `site_config`
  insert/update path the Peaks poster override admin UI already uses.

**Still blocked on**: the actual recipient email address(es) — Matt hasn't
provided the list yet (Q1 from the original plan doc). Ask him directly
before implementing.

## How to pull in the pending commits

A git bundle is included alongside this handoff (`handoff.bundle`). From a
local clone of this repo:

```
git fetch /path/to/handoff.bundle migration/google-appscript:incoming-handoff
git checkout migration/google-appscript
git merge incoming-handoff
git branch -d incoming-handoff
git push origin migration/google-appscript
```

Or if this `CLAUDE.md` and the rest of the tree already reflect the bundled
commits (i.e. you're working directly in the folder the bundle was built
from), there may be nothing to do but `git push`.

## Remaining phases (full detail in `docs/trailjournal-cowork-plan.md`)

Quick summary — see the plan doc for acceptance criteria and exact code:

- **Phase 2** — full-reflection email to a distribution list. Spec verified
  above; still needs Matt's recipient list before implementing.
- **Phase 3** — Fix the Google Sheet "full reflection" gap (missing headers
  silently drop fields on insert; add a readable/flattened column).
- **Phase 4** — Rebuild the admin dashboard (`admin/index.html`) on the Apps
  Script backend — it's currently still calling a Supabase client for some
  pages. This is the biggest remaining chunk of work.
- **Phase 5** — Close out: push everything, summarize for Matt, note
  ecosystem items intentionally left out of this plan (BIC URL-param reader,
  panel app migration, CoC widget paste-ins).

Also still outstanding from the original plan's question list:
- Google Sheet URL for the Trail Journal data (Q4 in the plan doc) — needed to
  actually verify/operate on the data side of Phases 2–4. Ask Matt.

## Working rules (carried over from the original plan)

- Branch per phase off `migration/google-appscript` (not `main`, per above)
  until the cutover happens.
- `node --check` every edited inline `<script>` block before considering an
  edit done (extract the script content to a temp `.js` file first).
- Test each phase on the live site (or a GitHub Pages preview of the working
  branch) before starting the next phase.
- Never commit Script Properties values. `APP_TOKEN` is already public in
  page source by design; `ANTHROPIC_API_KEY`/`ADMIN_PASSWORD` and other real
  secrets are not and must never end up in a commit.
