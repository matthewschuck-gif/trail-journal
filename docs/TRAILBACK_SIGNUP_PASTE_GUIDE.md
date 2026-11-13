# Trailback Panel sign-up: paste guide

Two Apps Script projects change: **Trail Journal** (the sign-up page, invites and case brief) and
**Climb-er** (writes each student's year pattern and strengths onto the case). Nothing is live
until you finish Part 1. Allow about 15 minutes.

How to copy a file from GitHub: open the link, click the **Copy raw file** button (two overlapping
squares, top right of the file), then in the Apps Script editor open the file of the same name,
select all (Ctrl+A), paste (Ctrl+V), and save (Ctrl+S).

## Part 1: Trail Journal

Open the Trail Journal Sheet, then Extensions > Apps Script.

### 1a. Two new files

| In Apps Script | Name it exactly | Copy from |
|---|---|---|
| + next to Files > **Script** | `11_PanelSignup` | [apps-script/11_PanelSignup.gs](https://github.com/matthewschuck-gif/trail-journal/blob/feature/panel-signup/apps-script/11_PanelSignup.gs) |
| + next to Files > **HTML** | `Signup` | [apps-script/Signup.html](https://github.com/matthewschuck-gif/trail-journal/blob/feature/panel-signup/apps-script/Signup.html) |

### 1b. Two existing files to replace (select all, paste over)

| File | Copy from | What changed |
|---|---|---|
| `03_Router` | [apps-script/03_Router.gs](https://github.com/matthewschuck-gif/trail-journal/blob/feature/panel-signup/apps-script/03_Router.gs) | Five new actions for the sign-up page |
| `08_Web` | [apps-script/08_Web.gs](https://github.com/matthewschuck-gif/trail-journal/blob/feature/panel-signup/apps-script/08_Web.gs) | The `signup` page, two new settings, one new menu item |

Save all four (Ctrl+S).

### 1c. Settings

1. Reload the Trail Journal Sheet. In the **Trail Journal** menu choose **Run setup**. This adds
   the two new rows to the Settings tab.
2. On the **Settings** tab, find `PANEL_ADMINS` and type the emails of everyone who may add cases
   and send briefs, separated by commas (you and Kelly). Left blank, only your account can.
3. `PANEL_QUESTION_SHEET_URL` already points at the v2 question sheet. Change it if the doc moves.

### 1d. Turn on the 3:00 brief

In the **Trail Journal** menu choose **Turn on Trailback case briefs (3:00 the day before)**.
Approve the permission prompt if one appears. This also makes the `panel_cases` and
`panel_signups` tabs.

### 1e. Redeploy the staff link

Apps Script editor > Deploy > Manage deployments > the **staff** deployment (Anyone within EASD) >
pencil icon > Version: **New version** > Deploy. Do the same for the student deployment so both
run the same code.

### 1f. Open the page

Until the GitHub branch is merged (Part 3), use the staff link with `?page=signup` on the end:

`https://script.google.com/a/macros/easdpa.org/s/AKfycbypRpxjHYGnuOlc7kObO19qqHCAdlVXV706DRxYz_nO0icOsOTh47iyybZu1p7sa2f4UA/exec?page=signup`

You should see "Trailback Panel sign-up", your email, and the Add a case form.

## Part 2: Climb-er

Open the Climb-er Sheet, then Extensions > Apps Script.

| In Apps Script | Copy from | What changed |
|---|---|---|
| + next to Files > **Script**, name it `Trailback` | [apps-script/Trailback.gs](https://github.com/matthewschuck-gif/student-360/blob/feature/trailback-briefs/apps-script/Trailback.gs) | New: builds the summary and writes it onto each case |
| `Code` (replace) | [apps-script/Code.gs](https://github.com/matthewschuck-gif/student-360/blob/feature/trailback-briefs/apps-script/Code.gs) | Three new settings, one new menu item |
| `SquadSync` (replace) | [apps-script/SquadSync.gs](https://github.com/matthewschuck-gif/student-360/blob/feature/trailback-briefs/apps-script/SquadSync.gs) | The hourly sync also fills the case briefs |

Then:

1. Save (Ctrl+S) and reload the Climb-er Sheet.
2. **Define a Climb-er** menu > **Run setup** (adds `TRAILBACK_BRIEFS`, `TRAILBACK_SHEET`,
   `TRAILBACK_DAYS` to Settings).
3. Make sure the hourly sync is on: **Turn on automatic squad sync**. The case briefs ride on it.
4. Deploy > Manage deployments > pencil > New version > Deploy.

## Part 3: The trailjournal.org/signup address

The short address comes from the `signup/` folder on GitHub. The site is served from the
`migration/google-appscript` branch, so merge `feature/panel-signup` into it (open a pull request
on GitHub and merge it). After that, `https://trailjournal.org/signup/` forwards to the staff link.

## Test it end to end (10 minutes)

1. On the sign-up page, add a case for a real student with their **800#**, dated tomorrow or later.
2. Click **Sit on this panel**. Check your inbox for the invite and open the attached calendar file.
3. In Climb-er: **Define a Climb-er** menu > **Fill Trailback case briefs now**. Reload the sign-up
   page; the case should show "Climb-er summary ready".
4. On the case, click **Send the brief now**. Check the email: squad and team, strengths on file,
   the pattern for the whole school year (counts, by month, and the records), then the Trail Journal.
5. Click **Step off this panel**. Check for the cancellation email.
6. Edit the test case and close sign-ups, or delete its row on the `panel_cases` tab.

## Things to know

- **Each panelist gets their own invite.** The district blocks the calendar API, so there is no
  single shared event. The invite email lists who else is on the panel.
- **Full student names show to any EASD staff member who opens the page.** To limit who can open
  any staff page, list emails in Settings > `STAFF_EMAILS`.
- **The 3:00 brief goes out the weekday before the panel.** It skips weekends but does not know
  days off. Use **Send the brief now** when a panel falls on the first day back.
- **The Climb-er summary needs the 800# on the case.** Without it the brief says the summary is
  not ready. It refreshes every hour while the case is open and upcoming.
- **The year pattern lists the newest 40 records** and counts the rest; the full history stays on
  the student's Climb-er page.
- **If something breaks,** paste the previous version of the file back: on GitHub, switch the
  branch picker to `migration/google-appscript` (Trail Journal) or `main` (Climb-er).
