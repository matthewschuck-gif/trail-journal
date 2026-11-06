# Every journal reaches the office and the Sheet (Nov 5, 2026)

## What was wrong
- After Submit, the page showed "done" and then waited for the AI summary. Only after the AI came back did it send the office email and write `reflection_text`, `goal` and `clash_flagged`.
- **Start New Journal** and **Close Out** reload the page, and Chromebooks sleep. If that happened during the wait, the email and those fields were lost. The exit auto-save is off on the "done" screen.
- Full-tier journals have the longest AI wait, so they were the ones lost. Detention journals usually finished, which is why only those emailed.
- The AI proxy itself always wrote `insight_level` and `ai_staff_summary` on the server.
- Partial journals only reached the Sheet through Ran Out of Time or the exit beacon.

## The fix
| Where | What |
|---|---|
| `02_AIProxy.gs` | When a `summary` finishes, `finishSubmission_` writes the report, goal and Clash flag and sends the office email on the server, even if the page is gone. |
| `04_FollowupEmails.gs` | The page's duplicate email is skipped; each email marks the row (`office_emailed`). Email attachments supported. |
| `10_Records.gs` (new) | `upsertReflection_` saves by session and never downgrades `submitted`. The safety net runs every 10 minutes: unsent rows get sent, and an `in_progress` row idle 60 minutes goes out as INCOMPLETE. `emailAllRecordsOnce` sends one email with every journal as a CSV, plus the missed ones written out, and fills the Sheet. `sendMissedRecords` sends them one by one. |
| `09_RecordText.gs` (new, generated) | The page's `buildEmailBody`, server-side. Regenerate it with `node tools/port-record-text.js` whenever the page version changes. |
| `01_Config.gs` | New `reflections` columns `office_emailed` and `last_saved_at` (added to existing Sheets automatically). |
| `03_Router.gs` | New `upsertReflection` action (no staff password needed, same as insert/update). |
| `08_Web.gs` | Menu: Turn on the record safety net, Email all journal records once…, Send missed journal records one by one… |
| `Journal.html`, `index.html` | Progress is saved to the Sheet at every screen (`in_progress`); every save goes through `upsertReflection`, falling back to insert/update on an old backend. |

Tests: `node tests/backend.test.js` (11) and `node tests/records.test.js` (6).

## Paste (with the move to Apps Script)
1. Paste `01_Config`, `02_AIProxy`, `03_Router`, `04_FollowupEmails`, `05_AdminAuth`, `07_ClashOfClasses`, `08_Web`.
2. Create and paste the new `09_RecordText` and `10_Records`.
3. Paste the page files: `Journal`, `Respond`, `Admin`, `Panel`, `Followup`, `TrailCss`, `TjBridge`.
4. Save, reload the Sheet, then Trail Journal › **Run setup**, then **Scramble gate passwords**.
5. Deploy › Manage deployments › pencil › **New version** (same /exec link).
6. Trail Journal › **Turn on the record safety net**.
7. Trail Journal › **Email all journal records once…** with the first day of school. This sends one real email to the Notifications recipients.
8. Test: submit a full-tier journal and click Start New Journal right away. The email should still arrive, and the row should have its report, goal and insight level.

Not recoverable: past journals that never reached the Sheet at all (they lived only on the device). The progress saves prevent that from now on.
