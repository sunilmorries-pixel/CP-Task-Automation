# CP Task Automation

Consolidates Tricog's Channel Partner (CP) field-service visit logs into one clean,
reliable tab, and mirrors it into the MASTER SERVICE workbook.

## What it does

CP visits are logged by hand across two tabs — "CP Tasks - 2025" and "CP Tasks 2026" —
in a workbook this project does not own (*Master Channel Task Tracker - 2025 & 2026*).
The two tabs aren't cleanly split by year, and neither is in a form the rest of the
reporting stack can consume directly.

`consolidateCpTasks` reads both tabs daily, applies four data-quality fixes (real parsed
dates, a corrected `TAT in days` that no longer goes negative on a visit with no closure
date, a trimmed `Ticket Status`, and a `Created month` derived from the date rather than
trusted from source text), and does a full rebuild of the `CP Master` tab in this
project's own workbook. It then mirrors that tab into MASTER SERVICE's `CP sheet` tab.

It is a full rebuild, not an upsert, because no column or combination of columns in the
source data is a reliable unique key. A run that reads zero rows from the source is
treated as a failure and leaves `CP Master` untouched, rather than risk blanking it out.

This project is independent of Tricog's ServiceWRK ticket automation
(`sunilmorries-pixel/Task-Tracker-automation`) — separate repo, separate Apps Script
project, no shared code. It only ever writes to its own `CP sheet` tab in the MASTER
SERVICE workbook; every other tab there belongs to that other project.

See [docs/specs/2026-09-04-cp-task-consolidation-design.md](docs/specs/2026-09-04-cp-task-consolidation-design.md)
for the full design.

> **`CP Master` is not safe to hand-edit.** Every run fully rebuilds its contents from
> the two source tabs; anything typed directly into `CP Master` is overwritten on the
> next run.

## Operating it

Deploy code changes with `npx clasp push --force`.

Functions runnable from the Apps Script editor's dropdown:

| Function | What it does |
|---|---|
| `consolidateCpTasks` | Reads both CP Tasks tabs, rebuilds `CP Master`, then mirrors it |
| `mirrorCpMaster` | Mirrors `CP Master` into MASTER SERVICE's `CP sheet` tab on its own |
| `installTriggers` | (Re)installs the daily trigger; clears existing ones first, safe to re-run |
| `removeAllTriggers_` | Rollback — stops all scheduled runs without touching any data (Apps Script's function dropdown may not list a function whose name ends in `_`, by convention treated as private — if it doesn't appear, delete the trigger directly from the Triggers page instead, clock icon in the left sidebar) |

Run the test suite with `npm test`.

CP Task Tracker's `Sync Log` tab records one row per run — timestamp, rows read from
each source tab, rows written, TAT values repaired, and status — and is the first place
to check after an alert email.

## Key resources

- CP Task Tracker (build target): https://docs.google.com/spreadsheets/d/1dWSDJUjafdnAWsbiXSwo9FlZXC_Qb3QqP8Psek6tPYI/edit
- Master Channel Task Tracker (read-only source): https://docs.google.com/spreadsheets/d/1kF1tDMvK7pmyqFG4ZTHEJ57mab802pw16a1ByBVmvGM/edit
- MASTER SERVICE (mirror target): https://docs.google.com/spreadsheets/d/16Q2q9R6GPBOBYVmvImRTZRp8g1kW-G6fio26XDJiULo/edit
- Apps Script project: https://script.google.com/d/1Mq15Zc6qaH-ctxQKmxHh2NvjzKfChurxgaFzQ9_4krWAQlFgiaXYGMBH/edit

## Project structure

```
src/                  Apps Script sources (pushed with clasp)
  Config.js           Resource IDs, schedule, tunables
  Alerts.js           Email (and optional Slack) error alerts
  SheetHelpers.js     Exact-name sheet resolution, grid-capacity helper
  CpConsolidate.js    Reads the two CP Tasks tabs, normalizes, rebuilds CP Master
  CpMirror.js         One-way mirror of CP Master into MASTER SERVICE
  Triggers.js         Trigger entry points and installer
  appsscript.json     Apps Script manifest (scopes, timezone)
  -- pure, unit-tested logic (no Apps Script globals) --
  DateParsing.js      CP's "22 December 2025" date format -> Date
  CpSchema.js         CP column schema and header-name mapping
  CpNormalize.js      The four CP data-quality fixes
  MirrorTargetGuard.js "Don't overwrite a foreign tab" mirror safety check
tests/                Jest tests for the pure modules
docs/                 Design spec and implementation plan
```

## Known limitations

- **`Customer ID` values like `"New spoke"` pass through unchanged.** A data-entry
  convention in the source, not an error; not usable as a join key against other data.
- **A CP visit is not deduplicated against the ServiceWRK ticket automation's Ticket
  Master**, even when the same Zoho ticket number appears in both. The two describe
  different things (a CP engineer visit vs. a ServiceWRK ticket lifecycle) tracked in
  two entirely separate automations.
- **`Sync Log`'s `Status` column reflects only the `CP Master` rebuild step.** If the
  subsequent mirror into MASTER SERVICE fails, `Sync Log` can still show `Success` for
  that run; the mirror failure is reported only by its own email alert (subject "CP
  mirror failed"), not reflected in `Sync Log`.
