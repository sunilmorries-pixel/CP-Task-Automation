# CP Task Consolidation — Design Spec

Date: 2026-09-04 (originally drafted in the ServiceWRK ticket automation repo; moved
here 2026-09-16 when the user decided CP automation should be its own repo/project)
Status: Approved, not yet implemented.

## Repository & Deployment

This project is standalone, deliberately separate from Tricog's existing ServiceWRK
ticket automation (repo `sunilmorries-pixel/Task-Tracker-automation`, Apps Script
project "Task tracker"):

- **This repo:** `sunilmorries-pixel/CP-Task-Automation`
- **Apps Script project:** "CP Task Automation" — `1Mq15Zc6qaH-ctxQKmxHh2NvjzKfChurxgaFzQ9_4krWAQlFgiaXYGMBH`
  (standalone, created 2026-09-16, no relation to the "Task tracker" project — `clasp`
  binds one repo to exactly one script ID, so two independently-pushed repos cannot
  safely share one Apps Script project)
- Runs as the same Google account as the ServiceWRK automation (the machine's `clasp`
  login is per-machine, not per-project), so it already has whatever access that account
  has to the Google Sheets referenced below — no new sharing should be needed, but this
  is worth confirming during implementation rather than assumed.

Nothing in this design depends on the ServiceWRK repo's code. Where this spec mentions
"the same shape as the ticket automation," that means the same *pattern* (a master tab,
a mirror, a sync log) reimplemented independently here — not shared code, since the two
projects have no shared codebase to share it from.

## Purpose

Tricog's Channel Partner (CP) field-service visits are logged by hand across two tabs
of a workbook owned outside this project — "CP Tasks - 2025" and "CP Tasks 2026" in
*Master Channel Task Tracker - 2025 & 2026*. The tabs are not cleanly split by year (the
2026 tab opens with 8 rows dated December 2025), so a reader has to open both to see the
current picture, and neither is in a form the rest of the reporting stack can consume.

This automation reads both tabs, unions them into one clean table, and writes the result
to a dedicated workbook plus a mirror tab in MASTER SERVICE.

It is a **read-only consumer** of the source workbook: nothing is ever written back to
*Master Channel Task Tracker*.

## Source Details

| | CP Tasks - 2025 | CP Tasks 2026 |
|---|---|---|
| Workbook | *Master Channel Task Tracker - 2025 & 2026* (`1kF1tDMvK7pmyqFG4ZTHEJ57mab802pw16a1ByBVmvGM`) | same |
| Owner | `dadapeer.z@tricog.com` | same |
| Rows sampled | 145 (reader-capped; true count unconfirmed, likely a few hundred) | 145 (reader-capped; stops at 17 Mar 2025 — also capped) |
| Date range observed | 1 Jan 2025 to 17 Mar 2025 (sample cutoff, not the tab's real end) | 22 Dec 2025 to 6 Jul 2026 |

Both tabs share an identical 17-column header, in the same order:

`CP Name`, `Engineer name`, `Created month`, `Created On`, `Zoho`, `Customer ID`,
`Customer Name`, `Customer Address`, `Service Type`, `Problem Description`,
`Ticket Type`, `Ticket Status`, `Closure comments`, `Closed On`, `Visit Type`,
`Allocated By`, `TAT in days`

No column in either tab is a reliable unique key: `Zoho` alone has 32 duplicate groups in
the 290-row sample plus 40 blank/`NA` values (one Zoho ticket legitimately covers several
centres); adding `Customer ID`, `Service Type`, and `Created On` still leaves 4 duplicate
groups. There is no keyed-upsert path available for this data.

## Target Resources

- **CP workbook (build target):** *CP Task Tracker* — `1dWSDJUjafdnAWsbiXSwo9FlZXC_Qb3QqP8Psek6tPYI`
  (created 2026-09-04 by `sunil.morries@tricog.com`, currently empty)
  - `CP Master` tab — the consolidated table
  - `Sync Log` tab — append-only CP run history
- **Mirror target:** MASTER SERVICE (`16Q2q9R6GPBOBYVmvImRTZRp8g1kW-G6fio26XDJiULo`),
  new `CP sheet` tab, alongside the existing `Service sheet` and the hand-curated `TOM`
  tab (neither touched by this work; both belong to the separate ServiceWRK automation)

## Architecture

```
Master Channel Task Tracker (read-only)
  |-- "CP Tasks - 2025" --\
  \-- "CP Tasks 2026"    --+--> consolidateCpTasks()
                                       |
                                       +--> CP Task Tracker
                                       |      |-- "CP Master" (full rebuild each run)
                                       |      \-- "Sync Log"  (append-only)
                                       |
                                       \--> mirrorCpMaster()
                                              \--> MASTER SERVICE "CP sheet"
```

`consolidateCpTasks` and `mirrorCpMaster` run back-to-back from one daily trigger, both
comfortably inside the 6-minute execution cap at this row count — no chunking, no resume
marker, no reason to split them onto separate triggers.

### Full rebuild, not upsert

Because no reliable key exists, every run replaces the entire contents of `CP Master`
from the two source tabs. This is not a compromise forced by time pressure — with no
key, merging would either duplicate rows or silently drop the ones a partial key
collides on. A full rebuild is the only version of this that cannot drift or corrupt.

Consequence, same as any mirror: `CP Master` is not safe to hand-edit. Anything typed
into it directly is overwritten on the next run.

## CP Master Schema

The 17 source columns, unchanged in position, plus two appended at the end (appending
never shifts an existing column, so nothing already written needs realignment if a
column is added later):

`CP Name`, `Engineer name`, `Created month`, `Created On`, `Zoho`, `Customer ID`,
`Customer Name`, `Customer Address`, `Service Type`, `Problem Description`,
`Ticket Type`, `Ticket Status`, `Closure comments`, `Closed On`, `Visit Type`,
`Allocated By`, `TAT in days`, `Source Tab`, `Created Year`

- `Source Tab` — which of the two source tabs the row came from, so a bad row can be
  traced back to its origin for a fix at source.
- `Created Year` — derived from the parsed `Created On`, not from the source tab name.
  Needed because the tabs are not year-clean (see Source Details above), so tab identity
  cannot be used for period grouping.

Sorted by `Created On` descending (newest first). Rows with an unparseable `Created On`
sort last, not dropped.

## Data Quality Fixes

Four columns carry values confirmed wrong or unreliable in the source data. Everything
else — including `Customer ID` values like `New spoke`, which are a data-entry
convention rather than an error — passes through verbatim.

1. **`Created On` / `Closed On` parsed into real dates.** Source format is text like
   `22 December 2025`. `DateParsing.js` provides `parseCpDate()` for this format.
2. **`TAT in days` recomputed** as `Closed On - Created On` in whole days, using the
   parsed dates, with a blank result (not a negative number) whenever `Closed On` is
   missing. This is the fix for the diagnosed bug: the source column is a spreadsheet
   formula that treats a blank `Closed On` as day-zero, producing values like `-46013`
   (the date serial for 22 Dec 2025) on 37 of 290 sampled rows — every one of them a
   `Cancelled` or `Failed` ticket, i.e. exactly the rows with no real closure date.
3. **`Ticket Status` trimmed** — source values carry trailing whitespace
   (confirmed against the raw sheet: `"Cancelled "` has a genuine trailing space, unlike
   `"Completed"`/`"Failed"`, which don't), which would otherwise split a single status
   into multiple groups under any `COUNTIF`/`QUERY` grouping.
4. **`Created month` recomputed** from the parsed `Created On` rather than copied from
   the source column, which disagrees with `Created On` on 5 of 290 sampled rows (e.g.
   month `April` against date `17 March 2025`).

All four are pure functions living in a new `src/CpNormalize.js`, unit-tested against
fixtures taken from the actual rows found during analysis (the `-46013` row, the
`April`/`17 March 2025` mismatch row, a `New spoke` customer ID) rather than invented
cases.

## Consolidation Logic

1. Open the source workbook by ID. List its actual tab names and resolve
   `CONFIG.CP_SOURCE_TAB_NAMES` against them — exact match, case-sensitive, no fallback,
   never `getSheetByName` (its case-insensitive matching has caused real data loss on a
   different project's sheet before — see the ServiceWRK ticket automation's own history
   for why this project resolves tab names explicitly instead of relying on that
   method's leniency). Missing tab aborts the run.
2. For each resolved tab, read the full used range in one `getValues()` call. Map the
   header row to the 17 expected columns by name, not position — tolerates the two
   tabs' columns being reordered relative to each other, and only aborts if an expected
   column name is missing entirely, not if an unrecognized extra column is present.
3. Normalize every row through `CpNormalize.js`, stamping `Source Tab`.
4. Concatenate both tabs' normalized rows and sort by parsed `Created On` descending.
5. Abort if the combined result is zero rows. A run that reads nothing is treated as a
   failure, not a valid empty state — otherwise a revoked share or a renamed tab would
   silently blank `CP Master` and propagate the blank into the MASTER SERVICE mirror.
6. Only once a complete, validated replacement set exists in memory: overwrite `CP
   Master`'s header and every row in one `setValues()` call, then trim any leftover rows
   from a larger previous run. A failure at any earlier step leaves the existing `CP
   Master` untouched.
7. Append one row to CP's `Sync Log`.
8. Call `mirrorCpMaster()`.

## Mirror into MASTER SERVICE

`mirrorCpMaster()` copies `CP Master` in full into the `CP sheet` tab of MASTER SERVICE.
At a few hundred rows this is one `getValues()` / one `setValues()` — no batching, no
resume marker.

A self-contained safety guard, `assertSafeMirrorTarget_(sheet, expectedColumns)`, lives
in this project's own `src/MirrorTargetGuard.js` and refuses to overwrite a tab whose
existing header doesn't match the expected schema — this is what protects `Service
sheet` and the hand-curated `TOM` tab in MASTER SERVICE from ever being overwritten by a
misconfigured `CONFIG.CP_MIRROR_SHEET_NAME`. It is written fresh in this project (there
is no shared codebase with the ServiceWRK automation to extract it from), but the logic
mirrors a proven pattern: an empty target tab is always safe; a tab whose header doesn't
match the expected columns throws before any write.

## Sync Log tab (CP Task Tracker)

Append-only:

| Timestamp | Rows Read (2025 tab) | Rows Read (2026 tab) | Rows Written | TAT Values Repaired | Status |
|---|---|---|---|---|---|

`Status` is `Success` or `Error: <short reason>`.

## Error Handling

| Failure | Behavior |
|---|---|
| A `CP_SOURCE_TAB_NAMES` entry doesn't resolve to an actual tab | Abort before any write, log `Error` to Sync Log, email alert |
| An expected column is missing from a source tab's header | Abort before any write, log `Error`, email alert |
| Source workbook can't be opened (access revoked) | Abort, log `Error`, email alert |
| Combined result across both tabs is zero rows | Treated as failure — abort, log `Error`, email alert, `CP Master` left untouched |
| `CP sheet` in MASTER SERVICE holds content whose header isn't the CP schema | `assertSafeMirrorTarget_` throws before any write |

Alerts go through a `sendErrorAlert_` function in this project's own `src/Alerts.js`
(same pattern as the ServiceWRK automation, reimplemented here) to
`sunil.morries@tricog.com`.

## Testing

Pure modules only — impure modules touching `SpreadsheetApp`/`PropertiesService` are
verified against the real sheets, not unit tested:

- **`CpSchema.js`** — header-to-object mapping is order-independent; a missing required
  column aborts; an unrecognized extra column is tolerated.
- **`CpNormalize.js`** — `parseCpDate` on real samples (`"22 December 2025"`,
  `"6 July 2026"`) and on garbage input; TAT recompute returns blank when `Closed On` is
  empty and a correct integer day count when both dates parse (regression case: the
  `-46013` row must come out blank, not negative); status trimming; month derivation
  against the real mismatch row.
- **`MirrorTargetGuard.js`** — allows an empty tab and a tab whose header matches;
  refuses a tab holding an unrecognized schema; checks only the overlapping width.
- **Sort comparator** — descending by parsed `Created On`; unparseable dates sort last.

## Out of Scope (explicitly deferred)

- Any write path back into *Master Channel Task Tracker* — this project only reads it
- A unique-key upsert for CP data — ruled out by the source data itself (see Source
  Details)
- Fixing the source `TAT in days` formula at its origin — flagged to raise with
  `dadapeer.z@tricog.com` separately; this design corrects the symptom downstream only
- Splitting `mirrorCpMaster` onto its own trigger window — not needed at this row count;
  revisit if CP volume ever approaches ServiceWRK's ticket-master scale
- Any integration with the ServiceWRK ticket automation's Apps Script project or Google
  Sheet — the two remain fully independent

## Known Limitations

- **Row counts from initial analysis were reader-capped at 145 per tab** and are not the
  true tab sizes; the true counts will be confirmed against the live sheet during
  implementation, but nothing in this design depends on the exact number.
- **`Customer ID` values like `"New spoke"` pass through unchanged.** No new customer ID
  is invented; a row like this is not usable as a join key against other data.
- **A CP visit that ServiceWRK also tracks under the same Zoho ticket is not
  deduplicated against Ticket Master.** The two describe different things (a CP engineer
  visit vs. a ServiceWRK ticket lifecycle) at different grains, tracked in two entirely
  separate automations; no attempt is made to reconcile them.
