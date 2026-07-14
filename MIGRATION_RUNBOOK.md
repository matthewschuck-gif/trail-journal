# Trail Journal → Google Sheets/Apps Script Migration Runbook

Everything is built and pushed. This is the checklist for your side: setting up the actual
Google Sheet + Apps Script project (I don't have write access to your Workspace), pointing the
frontend at it, testing, and cutting over. Nothing here requires touching code — it's all
clicking through Google's UI and pasting a few URLs.

## Where everything lives

- **`main`** — untouched, still the live Supabase-backed app exactly as it was.
- **`archive/supabase-original`** — a frozen copy of that same state, as an explicit rollback point.
- **`migration/google-appscript`** — the new Google-backed version. All 5 `.gs` backend files
  plus all 5 rewired frontend HTML files are here, fully committed.

## Part 1 — One-time Google setup

1. In your school Workspace Drive, create a new Google Sheet (any name — e.g. "Trail Journal Data").
2. Open it → **Extensions → Apps Script**.
3. Delete the default empty `Code.gs` file it creates.
4. Create 5 new script files (**File → New → Script file**) named exactly:
   `01_Config`, `02_AIProxy`, `03_Router`, `04_FollowupEmails`, `05_AdminAuth`
   and paste in the matching content from `apps-script/` in the `migration/google-appscript` branch.
5. **Project Settings (gear icon) → Script Properties**, add:
   | Property | Value |
   |---|---|
   | `ANTHROPIC_API_KEY` | your Claude API key (add this yourself, never share it in chat) |
   | `RESEND_API_KEY` | your existing Resend key (same one Supabase was using, or a new one) |
   | `FROM_EMAIL` | `noreply@trailjournal.org` (or whatever you prefer) |
   | `APP_TOKEN` | any random string you make up — this is the shared token the frontend sends on every request |
   | `ADMIN_PASSWORD` | any password you choose for staff access (replaces the old hardcoded password — this one lives only in Script Properties, never in frontend source) |
6. In the function dropdown at the top, select `setupSpreadsheet`, click **Run**. Approve the
   permission prompts. This creates all 6 tabs with correct headers.
7. **Go set sharing/protection on the `incident_reports` tab now** — it's the one with real
   student names, grades, and homerooms. Right-click the tab → Protect range → restrict to the
   same staff who'd have had admin access before.
8. **Deploy once** (Deploy → New deployment → type "Web app"): Execute as **Me**, Who has access
   **Anyone**. Copy this single URL — it's already wired into all 5 HTML files (see Part 2).

   **Why only one deployment:** an earlier version of this used a second, domain-restricted
   ("Anyone within your domain") deployment for staff pages. That doesn't actually work — Apps
   Script routes cross-origin requests to that deployment type through an extra Google
   auth-check hop that never returns CORS headers, so `trailjournal.org` could never read the
   response even when correctly signed in. Confirmed by real testing, not just a guess. The
   single public deployment plus the `ADMIN_PASSWORD` server-side check (in `05_AdminAuth.gs`)
   replaces it and actually works cross-origin.

## Part 2 — Point the frontend at your URLs

**Already done** — all 5 HTML files are wired to your single public deployment URL and your
`APP_TOKEN`, already pushed to the `migration/google-appscript` branch. Nothing to paste here
anymore. The only thing left in this step is adding the `ADMIN_PASSWORD` Script Property from
Part 1 — once that's set, staff pages will prompt for that password instead of Google sign-in.

## Part 3 — Test before touching production

**Do this on a separate test copy of the site first**, not on trailjournal.org directly. Easiest
way: enable GitHub Pages for the `migration/google-appscript` branch temporarily (Settings →
Pages → Branch), which gives you a `matthewschuck-gif.github.io/trail-journal/` URL to test
against without affecting the live domain.

Checklist:
- [ ] Main journal: complete a full Part 1 → 2 → 3 submission, confirm a row appears in the
      `reflections` tab and the AI reading/summary generates correctly.
- [ ] Responder form ("Your Side of the Trail"): submit, confirm a row in `responder_reflections`
      and the AI note generates.
- [ ] Staff section on the main journal page: enter the `ADMIN_PASSWORD` you set, confirm the
      gate unlocks (and confirm a wrong password shows "Incorrect password. Try again." and
      does not unlock).
- [ ] Admin dashboard: sign in, confirm dashboard stats, submissions table, content editor,
      pattern analysis, and parent letter generator all work.
- [ ] Trailback Panel: run through a full case, confirm `panel_sessions` row + `reflections`
      back-write (if you imported a journal session) both happen.
- [ ] Follow-up scheduler (both from the Panel and from the main journal's single-event button):
      confirm `followup_records` row + email with `.ics` attachment arrives.
- [ ] Confirm `incident_reports` sharing/protection is actually restricting access as expected.
- [ ] Confirm Clash of Classes still works normally (it's untouched, but worth a sanity check).

**About the single-event scheduler bug** (flagged inline in `index.html`): the button on the
main journal that's supposed to schedule one custom-dated follow-up actually sends parameters
the backend function has never read — it's always run the fixed 2-day/2-week/2-month sequence
instead, even in the original Supabase version. I ported it exactly as it currently behaves
rather than silently changing behavior during a migration. Decide if you want it actually fixed
to support a single custom date, and I can do that as a separate small change.

## Part 4 — Cut over

Once everything above checks out: merge `migration/google-appscript` into `main` (a normal PR
merge), turn GitHub Pages back to serving `main`, and trailjournal.org goes live on the new
backend. `archive/supabase-original` stays as your rollback point indefinitely.

## Housekeeping

- **GitHub token**: revoke it now if you haven't (Settings → Developer settings → Personal
  access tokens).
- **Archived Supabase project**: it's restored and active right now, holding Clash of Classes'
  data plus the old Trail Journal data as a reference. Free-tier projects auto-pause after 7 days
  idle and are *permanently deleted* after 90 days paused — so if you want the old Trail Journal
  data to survive as a long-term archive, either open the project every couple months, or ask me
  to set up a scheduled task that pings it periodically.
