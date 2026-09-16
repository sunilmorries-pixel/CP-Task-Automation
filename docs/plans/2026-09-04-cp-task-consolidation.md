# CP Task Consolidation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build the Apps Script functions that read the two CP Tasks tabs from the Channel Partner workbook, consolidate them into a `CP Master` tab in the *CP Task Tracker* workbook, and mirror the result into a `CP sheet` tab in MASTER SERVICE — per `docs/specs/2026-09-04-cp-task-consolidation-design.md`.

**Architecture:** A standalone Apps Script project (its own `clasp`-managed repo, no shared code with Tricog's separate ServiceWRK ticket automation). Pure, unit-tested logic modules (`CpSchema.js`, `CpNormalize.js`, `DateParsing.js`, `MirrorTargetGuard.js`) for header mapping, the four data-quality fixes, and the mirror safety check, plus GAS-only glue modules (`CpConsolidate.js`, `CpMirror.js`, `Alerts.js`, `SheetHelpers.js`, `Triggers.js`) for the actual reads/writes against `SpreadsheetApp`.

**Tech Stack:** Google Apps Script (V8 runtime), `clasp` (Google's Apps Script CLI), Node.js + Jest for the pure modules.

**Spec:** [docs/specs/2026-09-04-cp-task-consolidation-design.md](../specs/2026-09-04-cp-task-consolidation-design.md)

## Already done (project scaffold)

The following exist already and are not separate tasks — noted here so later steps aren't confused by files that predate them:

- `package.json` (name `cp-task-automation`, `jest` devDependency, `test` script)
- `tests/smoke.test.js`
- `.gitignore` (`node_modules/`, `.clasprc.json`, `.worktrees/`)
- `src/appsscript.json` — `timeZone: Asia/Kolkata`, `runtimeVersion: V8`, `oauthScopes`:
  `spreadsheets`, `script.send_mail`, `script.external_request`, `script.scriptapp`
  (no Gmail/Drive scopes — this project never reads mail or file attachments)
- `.clasp.json` — `scriptId: 1Mq15Zc6qaH-ctxQKmxHh2NvjzKfChurxgaFzQ9_4krWAQlFgiaXYGMBH`,
  `rootDir: src` (standalone Apps Script project "CP Task Automation", created via
  `clasp create`, no relation to the ServiceWRK ticket automation's "Task tracker"
  project)

## Global Constraints

- CP source workbook ID `1kF1tDMvK7pmyqFG4ZTHEJ57mab802pw16a1ByBVmvGM` ("Master Channel Task Tracker - 2025 & 2026", owned by `dadapeer.z@tricog.com`) — pre-existing, read-only, never write to it.
- CP source tab names are `'CP Tasks - 2025'` and `'CP Tasks 2026'` — exact match, case-sensitive. Never use `getSheetByName` to resolve them; resolve by comparing `sheet.getName() ===` against the configured name instead (see the design spec's Consolidation Logic section for why).
- CP workbook ID `1dWSDJUjafdnAWsbiXSwo9FlZXC_Qb3QqP8Psek6tPYI` ("CP Task Tracker", already created, currently empty) — the build target. Do not create a new workbook.
- MASTER SERVICE ID `16Q2q9R6GPBOBYVmvImRTZRp8g1kW-G6fio26XDJiULo` — existing workbook belonging to the separate ServiceWRK automation, already used by its own ticket mirror. This project only ever touches its `CP sheet` tab. Never touch `Service sheet` or `TOM`.
- Tab names: `CP Master` (in CP Task Tracker), `Sync Log` (in CP Task Tracker), `CP sheet` (in MASTER SERVICE).
- Full rebuild only, never a keyed upsert. The CP data has no reliable unique key (verified during design: `Zoho` + `Customer ID` + `Service Type` + `Created On` still leaves duplicate groups).
- A run that reads zero rows across both source tabs is a failure, not a valid empty state — abort before writing, alert, and leave `CP Master` untouched.
- Alert emails go to `sunil.morries@tricog.com`.
- This project shares no code and no Apps Script project with the ServiceWRK ticket automation — every module it needs (date parsing, the mirror safety guard, error alerting, grid-capacity helpers) is implemented fresh here, even where an equivalent already exists in that other, separate repo.
- Every pure logic file (no GAS globals) must remain requirable under plain Node — do not reference `SpreadsheetApp`, `MailApp`, `UrlFetchApp`, `ScriptApp`, or `PropertiesService` from `DateParsing.js`, `CpSchema.js`, `CpNormalize.js`, or `MirrorTargetGuard.js`.

---

### Task 1: Config.js — central configuration

**Files:**
- Create: `src/Config.js`
- Test: `tests/Config.test.js`

**Interfaces:**
- Consumes: nothing
- Produces: global `CONFIG` object (GAS) / `{ CONFIG }` (Node), consumed by every later GAS-only task via the bare `CONFIG` identifier

- [ ] **Step 1: Write the failing test**

`tests/Config.test.js`:

```js
const { CONFIG } = require('../src/Config');

test('CP source workbook and tabs point at the approved Channel Partner tracker', () => {
  expect(CONFIG.CP_SOURCE_SHEET_ID).toBe('1kF1tDMvK7pmyqFG4ZTHEJ57mab802pw16a1ByBVmvGM');
  expect(CONFIG.CP_SOURCE_TAB_NAMES).toEqual(['CP Tasks - 2025', 'CP Tasks 2026']);
});

test('CP workbook and mirror target match the approved design spec', () => {
  expect(CONFIG.CP_WORKBOOK_SHEET_ID).toBe('1dWSDJUjafdnAWsbiXSwo9FlZXC_Qb3QqP8Psek6tPYI');
  expect(CONFIG.CP_MASTER_SHEET_NAME).toBe('CP Master');
  expect(CONFIG.CP_SYNC_LOG_SHEET_NAME).toBe('Sync Log');
  expect(CONFIG.MASTER_SERVICE_SHEET_ID).toBe('16Q2q9R6GPBOBYVmvImRTZRp8g1kW-G6fio26XDJiULo');
  expect(CONFIG.CP_MIRROR_SHEET_NAME).toBe('CP sheet');
  expect(CONFIG.CP_TRIGGER_TIMES).toHaveLength(1);
});

test('the CP workbook, source workbook, and mirror target are three distinct resources', () => {
  expect(CONFIG.CP_WORKBOOK_SHEET_ID).not.toBe(CONFIG.CP_SOURCE_SHEET_ID);
  expect(CONFIG.CP_WORKBOOK_SHEET_ID).not.toBe(CONFIG.MASTER_SERVICE_SHEET_ID);
  expect(CONFIG.CP_SOURCE_SHEET_ID).not.toBe(CONFIG.MASTER_SERVICE_SHEET_ID);
});

test('the CP mirror tab name cannot collide with CP Master or Sync Log, case-insensitively', () => {
  const mirror = CONFIG.CP_MIRROR_SHEET_NAME.toLowerCase();
  expect(mirror).not.toBe(CONFIG.CP_MASTER_SHEET_NAME.toLowerCase());
  expect(mirror).not.toBe(CONFIG.CP_SYNC_LOG_SHEET_NAME.toLowerCase());
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx jest tests/Config.test.js`
Expected: FAIL — `Cannot find module '../src/Config'`.

- [ ] **Step 3: Write `src/Config.js`**

```js
var CONFIG = {
  // Source: read-only, owned outside this project.
  CP_SOURCE_SHEET_ID: '1kF1tDMvK7pmyqFG4ZTHEJ57mab802pw16a1ByBVmvGM',
  CP_SOURCE_TAB_NAMES: ['CP Tasks - 2025', 'CP Tasks 2026'],

  // Build target: this project's own workbook.
  CP_WORKBOOK_SHEET_ID: '1dWSDJUjafdnAWsbiXSwo9FlZXC_Qb3QqP8Psek6tPYI',
  CP_MASTER_SHEET_NAME: 'CP Master',
  CP_SYNC_LOG_SHEET_NAME: 'Sync Log',

  // Mirror target: MASTER SERVICE, belonging to the separate ServiceWRK ticket
  // automation. Only CP_MIRROR_SHEET_NAME is ever written; every other tab in
  // that workbook belongs to that other project and must never be touched.
  MASTER_SERVICE_SHEET_ID: '16Q2q9R6GPBOBYVmvImRTZRp8g1kW-G6fio26XDJiULo',
  CP_MIRROR_SHEET_NAME: 'CP sheet',

  CP_TRIGGER_TIMES: [[8, 30]],
  ALERT_EMAIL: 'sunil.morries@tricog.com',
  SLACK_WEBHOOK_URL: ''
};

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { CONFIG: CONFIG };
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `npx jest tests/Config.test.js`
Expected: PASS, 4 tests.

- [ ] **Step 5: Commit**

```bash
git add src/Config.js tests/Config.test.js
git commit -m "feat: add central Config.js" -m "Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 2: DateParsing.js — CP's text date format

**Files:**
- Create: `src/DateParsing.js`
- Test: `tests/DateParsing.test.js`

**Interfaces:**
- Consumes: nothing
- Produces: `parseCpDate(value)` — `(string) -> Date|null`

- [ ] **Step 1: Write the failing tests**

`tests/DateParsing.test.js`:

```js
const { parseCpDate } = require('../src/DateParsing');

test('parses a CP-format date with a two-digit day', () => {
  const result = parseCpDate('22 December 2025');
  expect(result.getFullYear()).toBe(2025);
  expect(result.getMonth()).toBe(11);
  expect(result.getDate()).toBe(22);
});

test('parses a CP-format date with a one-digit day', () => {
  const result = parseCpDate('6 July 2026');
  expect(result.getFullYear()).toBe(2026);
  expect(result.getMonth()).toBe(6);
  expect(result.getDate()).toBe(6);
});

test('parses every month name', () => {
  expect(parseCpDate('1 January 2026').getMonth()).toBe(0);
  expect(parseCpDate('1 February 2026').getMonth()).toBe(1);
  expect(parseCpDate('1 March 2026').getMonth()).toBe(2);
  expect(parseCpDate('1 April 2026').getMonth()).toBe(3);
  expect(parseCpDate('1 May 2026').getMonth()).toBe(4);
  expect(parseCpDate('1 June 2026').getMonth()).toBe(5);
  expect(parseCpDate('1 August 2026').getMonth()).toBe(7);
  expect(parseCpDate('1 September 2026').getMonth()).toBe(8);
  expect(parseCpDate('1 October 2026').getMonth()).toBe(9);
  expect(parseCpDate('1 November 2026').getMonth()).toBe(10);
  expect(parseCpDate('1 December 2026').getMonth()).toBe(11);
});

test('rejects a calendar-invalid date rather than letting it roll over', () => {
  // JS Date silently turns "31 April" into "1 May" if not guarded against.
  expect(parseCpDate('31 April 2026')).toBeNull();
});

test('returns null for blank or unrecognized values', () => {
  expect(parseCpDate('')).toBeNull();
  expect(parseCpDate(null)).toBeNull();
  expect(parseCpDate('not a date')).toBeNull();
  expect(parseCpDate('22-12-2025')).toBeNull();
});
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npx jest tests/DateParsing.test.js`
Expected: FAIL — `Cannot find module '../src/DateParsing'`.

- [ ] **Step 3: Write `src/DateParsing.js`**

```js
// CP source rows write dates as free text like "22 December 2025".
var CP_MONTH_NAMES = [
  'january', 'february', 'march', 'april', 'may', 'june',
  'july', 'august', 'september', 'october', 'november', 'december'
];

function parseCpDate(value) {
  if (!value) return null;
  var str = String(value).trim();
  var match = str.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/);
  if (!match) return null;
  var day = Number(match[1]);
  var monthIndex = CP_MONTH_NAMES.indexOf(match[2].toLowerCase());
  var year = Number(match[3]);
  if (monthIndex === -1) return null;
  var date = new Date(year, monthIndex, day);
  // Date silently rolls an invalid day into the next month (e.g. "31 April"
  // becomes "1 May") instead of erroring, so confirm it landed where asked.
  if (date.getMonth() !== monthIndex || date.getDate() !== day) return null;
  return date;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { parseCpDate: parseCpDate };
}
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `npx jest tests/DateParsing.test.js`
Expected: PASS, 6 tests.

- [ ] **Step 5: Commit**

```bash
git add src/DateParsing.js tests/DateParsing.test.js
git commit -m "feat: add CP text-date parsing" -m "Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 3: CpSchema.js — CP column schema and header mapping

**Files:**
- Create: `src/CpSchema.js`
- Test: `tests/CpSchema.test.js`

**Interfaces:**
- Consumes: nothing
- Produces: `CP_SOURCE_COLUMNS` (array of 17 strings), `CP_DERIVED_COLUMNS` (array of 2 strings), `CP_MASTER_COLUMNS` (array of 19 strings), `findMissingCpColumns(headers)` — `(array) -> array` of missing column names, `mapCpRow(headers, rowValues)` — `(array, array) -> object` keyed by `CP_SOURCE_COLUMNS` names

- [ ] **Step 1: Write the failing tests**

`tests/CpSchema.test.js`:

```js
const { CP_SOURCE_COLUMNS, CP_DERIVED_COLUMNS, CP_MASTER_COLUMNS, findMissingCpColumns, mapCpRow } = require('../src/CpSchema');

test('CP_SOURCE_COLUMNS matches the real header of both CP Tasks tabs', () => {
  expect(CP_SOURCE_COLUMNS).toEqual([
    'CP Name', 'Engineer name', 'Created month', 'Created On', 'Zoho', 'Customer ID',
    'Customer Name', 'Customer Address', 'Service Type', 'Problem Description',
    'Ticket Type', 'Ticket Status', 'Closure comments', 'Closed On', 'Visit Type',
    'Allocated By', 'TAT in days'
  ]);
});

test('CP_MASTER_COLUMNS appends the derived columns after the 17 source columns', () => {
  expect(CP_MASTER_COLUMNS.length).toBe(19);
  expect(CP_MASTER_COLUMNS.slice(0, 17)).toEqual(CP_SOURCE_COLUMNS);
  expect(CP_MASTER_COLUMNS.slice(17)).toEqual(['Source Tab', 'Created Year']);
  expect(CP_DERIVED_COLUMNS).toEqual(['Source Tab', 'Created Year']);
});

test('findMissingCpColumns returns nothing when every required column is present', () => {
  expect(findMissingCpColumns(CP_SOURCE_COLUMNS)).toEqual([]);
});

test('findMissingCpColumns tolerates the two tabs having columns in a different order', () => {
  const reordered = CP_SOURCE_COLUMNS.slice().reverse();
  expect(findMissingCpColumns(reordered)).toEqual([]);
});

test('findMissingCpColumns reports a missing required column by name', () => {
  const withoutZoho = CP_SOURCE_COLUMNS.filter((c) => c !== 'Zoho');
  expect(findMissingCpColumns(withoutZoho)).toEqual(['Zoho']);
});

test('findMissingCpColumns tolerates an unrecognized extra column', () => {
  const withExtra = CP_SOURCE_COLUMNS.concat(['Some New Column']);
  expect(findMissingCpColumns(withExtra)).toEqual([]);
});

test('mapCpRow maps by header name, not position', () => {
  const headers = ['Zoho', 'CP Name'];
  const row = mapCpRow(headers, ['#125704', 'S S Medical System']);
  expect(row['CP Name']).toBe('S S Medical System');
  expect(row['Zoho']).toBe('#125704');
});

test('mapCpRow fills an absent column with an empty string rather than undefined', () => {
  const row = mapCpRow(['CP Name'], ['S S Medical System']);
  expect(row['Zoho']).toBe('');
  expect(row['TAT in days']).toBe('');
});

test('mapCpRow only reads columns it knows about, ignoring an unrecognized extra one', () => {
  const headers = ['CP Name', 'Some New Column'];
  const row = mapCpRow(headers, ['S S Medical System', 'ignore me']);
  expect(row['CP Name']).toBe('S S Medical System');
  expect(row['Some New Column']).toBeUndefined();
});
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npx jest tests/CpSchema.test.js`
Expected: FAIL — `Cannot find module '../src/CpSchema'`.

- [ ] **Step 3: Write `src/CpSchema.js`**

```js
var CP_SOURCE_COLUMNS = [
  'CP Name', 'Engineer name', 'Created month', 'Created On', 'Zoho', 'Customer ID',
  'Customer Name', 'Customer Address', 'Service Type', 'Problem Description',
  'Ticket Type', 'Ticket Status', 'Closure comments', 'Closed On', 'Visit Type',
  'Allocated By', 'TAT in days'
];

var CP_DERIVED_COLUMNS = ['Source Tab', 'Created Year'];

var CP_MASTER_COLUMNS = CP_SOURCE_COLUMNS.concat(CP_DERIVED_COLUMNS);

function findMissingCpColumns(headers) {
  var present = headers.map(function (h) { return String(h || '').trim(); });
  return CP_SOURCE_COLUMNS.filter(function (col) {
    return present.indexOf(col) === -1;
  });
}

function mapCpRow(headers, rowValues) {
  var trimmedHeaders = headers.map(function (h) { return String(h || '').trim(); });
  var row = {};
  CP_SOURCE_COLUMNS.forEach(function (col) {
    var index = trimmedHeaders.indexOf(col);
    row[col] = index === -1 ? '' : rowValues[index];
  });
  return row;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    CP_SOURCE_COLUMNS: CP_SOURCE_COLUMNS,
    CP_DERIVED_COLUMNS: CP_DERIVED_COLUMNS,
    CP_MASTER_COLUMNS: CP_MASTER_COLUMNS,
    findMissingCpColumns: findMissingCpColumns,
    mapCpRow: mapCpRow
  };
}
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `npx jest tests/CpSchema.test.js`
Expected: PASS, 9 tests.

- [ ] **Step 5: Commit**

```bash
git add src/CpSchema.js tests/CpSchema.test.js
git commit -m "feat: add CpSchema.js for CP column schema and header mapping" -m "Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 4: CpNormalize.js — the four data-quality fixes

**Files:**
- Create: `src/CpNormalize.js`
- Test: `tests/CpNormalize.test.js`

**Interfaces:**
- Consumes: `parseCpDate` (`src/DateParsing.js`, Task 2)
- Produces: `computeCpTatDays(createdDate, closedDate)` — `(Date|null, Date|null) -> number|''`, `deriveCreatedMonthName(createdDate)` — `(Date|null) -> string`, `deriveCreatedYear(createdDate)` — `(Date|null) -> number|''`, `normalizeCpRow(rawRow, sourceTab)` — `(object, string) -> object` keyed by `CP_MASTER_COLUMNS` names, `compareCpRowsByCreatedOnDesc(a, b)` — a `.sort()` comparator over `normalizeCpRow` output

All test fixtures below are real rows found during design analysis of the CP source workbook (not invented), reproduced verbatim.

- [ ] **Step 1: Write the failing tests**

`tests/CpNormalize.test.js`:

```js
const {
  computeCpTatDays,
  deriveCreatedMonthName,
  deriveCreatedYear,
  normalizeCpRow,
  compareCpRowsByCreatedOnDesc
} = require('../src/CpNormalize');

// --- computeCpTatDays -------------------------------------------------

test('returns blank when there is no closed date (cancelled/failed visits)', () => {
  expect(computeCpTatDays(new Date(2025, 11, 22), null)).toBe('');
});

test('returns blank when there is no created date', () => {
  expect(computeCpTatDays(null, new Date(2025, 11, 22))).toBe('');
});

test('computes a same-day turnaround as 0', () => {
  expect(computeCpTatDays(new Date(2025, 2, 17), new Date(2025, 2, 17))).toBe(0);
});

test('computes a multi-day turnaround, including across a month/year boundary', () => {
  // Real row: created 23 December 2025, closed 2 January 2026, source TAT
  // column agreed on 10 — this is the case where recomputing and the
  // (correct, non-broken) source value should match.
  expect(computeCpTatDays(new Date(2025, 11, 23), new Date(2026, 0, 2))).toBe(10);
});

// --- deriveCreatedMonthName / deriveCreatedYear ------------------------

test('derives the month name from the date, not from any source text', () => {
  expect(deriveCreatedMonthName(new Date(2025, 2, 17))).toBe('March');
});

test('derives the year from the date', () => {
  expect(deriveCreatedYear(new Date(2025, 2, 17))).toBe(2025);
});

test('month and year derivation return blank for a missing date', () => {
  expect(deriveCreatedMonthName(null)).toBe('');
  expect(deriveCreatedYear(null)).toBe('');
});

// --- normalizeCpRow: real fixture rows ---------------------------------

test('a cancelled row with the broken -46013 TAT gets a blank TAT, not the broken number', () => {
  // Real row, first data row of "CP Tasks 2026": Ticket Status source text is
  // "Cancelled " (one genuine trailing space, confirmed against the raw
  // sheet — "Completed"/"Failed" carry none), Closed On is blank, and the
  // source TAT column reads -46013 (the date serial for 22 December 2025
  // minus zero, i.e. the source formula treating a blank Closed On as
  // day-zero).
  const raw = {
    'CP Name': 'S S Medical System',
    'Engineer name': 'Vivek',
    'Created month': 'December',
    'Created On': '22 December 2025',
    'Zoho': '#125704',
    'Customer ID': '40756',
    'Customer Name': 'CP Hospital',
    'Customer Address': 'Piprajatampur Kubersthan, null, Kushinagar - 274304',
    'Service Type': 'Identification of the issue',
    'Problem Description': 'VCARDIA',
    'Ticket Type': 'CORE_SERVICE_JOB',
    'Ticket Status': 'Cancelled ',
    'Closure comments': '',
    'Closed On': '',
    'Visit Type': '',
    'Allocated By': 'Dadapeer',
    'TAT in days': '-46013'
  };
  const row = normalizeCpRow(raw, 'CP Tasks 2026');
  expect(row['TAT in days']).toBe('');
  expect(row['Ticket Status']).toBe('Cancelled');
  expect(row['Created On'].getFullYear()).toBe(2025);
  expect(row['Created On'].getMonth()).toBe(11);
  expect(row['Created On'].getDate()).toBe(22);
  expect(row['Created month']).toBe('December');
  expect(row['Created Year']).toBe(2025);
  expect(row['Source Tab']).toBe('CP Tasks 2026');
  expect(row['CP Name']).toBe('S S Medical System');
  expect(row['Customer ID']).toBe('40756');
});

test('a row whose source month text disagrees with its date gets the date-derived month', () => {
  // Real row, "CP Tasks - 2025": source "Created month" reads April, but
  // "Created On" reads 17 March 2025 — a genuine mismatch found during
  // analysis.
  const raw = {
    'CP Name': 'Shree Sai Healthcare',
    'Engineer name': 'Logesh',
    'Created month': 'April',
    'Created On': '17 March 2025',
    'Zoho': '#101640',
    'Customer ID': '41815',
    'Customer Name': 'Palanisamy Trumacare & Hospital',
    'Customer Address': 'Athani Road, komanapalayam, Sathyamangalam , Erode - 638401',
    'Service Type': 'Identification of the issue',
    'Problem Description': 'VCARDIA',
    'Ticket Type': 'CORE_SERVICE_JOB',
    'Ticket Status': 'Completed',
    'Closure comments': 'visited the center and found OTG cable issue so swapped the OTG cable after collecting the payment',
    'Closed On': '17 March 2025',
    'Visit Type': 'CENTER_VISIT',
    'Allocated By': 'Saidha Rao',
    'TAT in days': '0'
  };
  const row = normalizeCpRow(raw, 'CP Tasks - 2025');
  expect(row['Created month']).toBe('March');
  expect(row['Created Year']).toBe(2025);
  expect(row['TAT in days']).toBe(0);
});

test('a non-numeric Customer ID sentinel passes through unchanged', () => {
  // Real row: Customer ID is the text "New spoke", a data-entry convention,
  // not an error — normalizeCpRow must not invent or blank it out.
  const raw = {
    'CP Name': 'Hayana Enterprises',
    'Engineer name': 'Sumit',
    'Created month': 'February',
    'Created On': '7 February 2026',
    'Zoho': '#129558',
    'Customer ID': 'New spoke',
    'Customer Name': 'Health Central',
    'Customer Address': 'Kankeshwari Devi Mandir, 84/2, Pardesi Pura Main Rd, Near Readymade Complex, Pardesipura, Indore, Madhya Pradesh 452010',
    'Service Type': 'Installation',
    'Problem Description': 'VCARDIA',
    'Ticket Type': 'CORE_SERVICE_JOB',
    'Ticket Status': 'Completed',
    'Closure comments': 'Installation completed',
    'Closed On': '9 February 2026',
    'Visit Type': 'CENTER_VISIT',
    'Allocated By': 'Dadapeer',
    'TAT in days': '2'
  };
  const row = normalizeCpRow(raw, 'CP Tasks 2026');
  expect(row['Customer ID']).toBe('New spoke');
  expect(row['TAT in days']).toBe(2);
});

test('an unparseable Created On is kept visible as the original text, not dropped', () => {
  const raw = {
    'CP Name': 'X', 'Engineer name': 'Y', 'Created month': 'Unknown',
    'Created On': 'not a date', 'Zoho': 'NA', 'Customer ID': '1',
    'Customer Name': 'Z', 'Customer Address': '', 'Service Type': '',
    'Problem Description': '', 'Ticket Type': '', 'Ticket Status': 'Completed',
    'Closure comments': '', 'Closed On': '', 'Visit Type': '',
    'Allocated By': '', 'TAT in days': ''
  };
  const row = normalizeCpRow(raw, 'CP Tasks 2026');
  expect(row['Created On']).toBe('not a date');
  expect(row['Created month']).toBe('');
  expect(row['Created Year']).toBe('');
});

// --- compareCpRowsByCreatedOnDesc --------------------------------------

test('sorts by Created On descending, newest first', () => {
  const older = { 'Created On': new Date(2025, 0, 1) };
  const newer = { 'Created On': new Date(2026, 0, 1) };
  expect([older, newer].sort(compareCpRowsByCreatedOnDesc)).toEqual([newer, older]);
});

test('sorts rows with an unparseable Created On to the end, not dropped', () => {
  const dated = { 'Created On': new Date(2025, 0, 1) };
  const undated = { 'Created On': 'not a date' };
  expect([undated, dated].sort(compareCpRowsByCreatedOnDesc)).toEqual([dated, undated]);
  expect([dated, undated].sort(compareCpRowsByCreatedOnDesc)).toEqual([dated, undated]);
});
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npx jest tests/CpNormalize.test.js`
Expected: FAIL — `Cannot find module '../src/CpNormalize'`.

- [ ] **Step 3: Write `src/CpNormalize.js`**

```js
var DateParsingLib = (typeof require !== 'undefined')
  ? require('./DateParsing')
  : { parseCpDate: parseCpDate };

var CP_MONTH_DISPLAY_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'
];

// Fix 2 (design spec): TAT in days recomputed from the parsed dates rather
// than trusted from the source column. The source formula treats a blank
// Closed On as day-zero, producing large negative "TAT" values (e.g.
// -46013) on every Cancelled/Failed row — exactly the rows with no real
// closure date. Blank is the correct result there, not a number.
function computeCpTatDays(createdDate, closedDate) {
  if (!(createdDate instanceof Date) || !(closedDate instanceof Date)) {
    return '';
  }
  var msPerDay = 24 * 60 * 60 * 1000;
  return Math.round((closedDate.getTime() - createdDate.getTime()) / msPerDay);
}

// Fix 4: Created month derived from the parsed date. The source "Created
// month" column disagrees with "Created On" on some rows (found during
// analysis), so it cannot be trusted verbatim.
function deriveCreatedMonthName(createdDate) {
  if (!(createdDate instanceof Date)) return '';
  return CP_MONTH_DISPLAY_NAMES[createdDate.getMonth()];
}

// Derived column: which calendar year a row belongs to, from the date
// itself rather than the source tab's name — the tabs are not year-clean
// (the 2026 tab opens with December-2025 rows).
function deriveCreatedYear(createdDate) {
  if (!(createdDate instanceof Date)) return '';
  return createdDate.getFullYear();
}

// One CpSchema.mapCpRow() object + which tab it came from -> the final
// 19-column CP_MASTER_COLUMNS-keyed row. Everything not named in one of the
// four data-quality fixes (fix 1: dates; fix 2: TAT; fix 3: status; fix 4:
// month) passes through verbatim — including sentinel values like a
// Customer ID of "New spoke", which are a data-entry convention, not an
// error to correct.
function normalizeCpRow(rawRow, sourceTab) {
  var createdDate = DateParsingLib.parseCpDate(rawRow['Created On']);
  var closedDate = DateParsingLib.parseCpDate(rawRow['Closed On']);
  return {
    'CP Name': rawRow['CP Name'],
    'Engineer name': rawRow['Engineer name'],
    'Created month': deriveCreatedMonthName(createdDate),
    // Fix 1: a parsed Date when possible, so the sheet gets a real, sortable
    // date column instead of free text. Falls back to the original text when
    // unparseable, so a bad row is visible rather than silently blanked.
    'Created On': createdDate || rawRow['Created On'],
    'Zoho': rawRow['Zoho'],
    'Customer ID': rawRow['Customer ID'],
    'Customer Name': rawRow['Customer Name'],
    'Customer Address': rawRow['Customer Address'],
    'Service Type': rawRow['Service Type'],
    'Problem Description': rawRow['Problem Description'],
    'Ticket Type': rawRow['Ticket Type'],
    // Fix 3: trimmed. The source has been observed with a real trailing
    // space on "Cancelled " specifically, which would otherwise sit apart
    // from a clean "Cancelled" typed into any future row or formula.
    'Ticket Status': String(rawRow['Ticket Status'] || '').trim(),
    'Closure comments': rawRow['Closure comments'],
    'Closed On': closedDate || rawRow['Closed On'],
    'Visit Type': rawRow['Visit Type'],
    'Allocated By': rawRow['Allocated By'],
    'TAT in days': computeCpTatDays(createdDate, closedDate),
    'Source Tab': sourceTab,
    'Created Year': deriveCreatedYear(createdDate)
  };
}

// Newest first. A row whose Created On failed to parse (still a string, not
// a Date, per normalizeCpRow's fallback) sorts after every dated row rather
// than being dropped or sorting arbitrarily.
function compareCpRowsByCreatedOnDesc(a, b) {
  var aDate = a['Created On'] instanceof Date ? a['Created On'] : null;
  var bDate = b['Created On'] instanceof Date ? b['Created On'] : null;
  if (aDate && bDate) return bDate.getTime() - aDate.getTime();
  if (aDate && !bDate) return -1;
  if (!aDate && bDate) return 1;
  return 0;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    computeCpTatDays: computeCpTatDays,
    deriveCreatedMonthName: deriveCreatedMonthName,
    deriveCreatedYear: deriveCreatedYear,
    normalizeCpRow: normalizeCpRow,
    compareCpRowsByCreatedOnDesc: compareCpRowsByCreatedOnDesc
  };
}
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `npx jest tests/CpNormalize.test.js`
Expected: PASS, 13 tests.

- [ ] **Step 5: Commit**

```bash
git add src/CpNormalize.js tests/CpNormalize.test.js
git commit -m "feat: add CpNormalize.js with the four CP data-quality fixes" -m "Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 5: MirrorTargetGuard.js — the mirror safety guard

**Files:**
- Create: `src/MirrorTargetGuard.js`
- Test: `tests/MirrorTargetGuard.test.js`

**Interfaces:**
- Consumes: nothing
- Produces: `assertSafeMirrorTarget_(sheet, expectedColumns)` — `(sheet, array) -> void`, throws if unsafe. `sheet` needs only `getLastRow()`, `getLastColumn()`, `getName()`, and `getRange(r, c, nr, nc).getValues()` — no other Apps Script surface, which is what makes it fakeable in Jest without touching `SpreadsheetApp`.

- [ ] **Step 1: Write the failing tests**

`tests/MirrorTargetGuard.test.js`:

```js
const { assertSafeMirrorTarget_ } = require('../src/MirrorTargetGuard');

function fakeSheet(name, header) {
  return {
    getName: function () { return name; },
    getLastRow: function () { return header ? 5 : 0; },
    getLastColumn: function () { return header ? header.length : 0; },
    getRange: function () {
      return { getValues: function () { return [header]; } };
    }
  };
}

test('allows a completely empty tab', () => {
  expect(() => assertSafeMirrorTarget_(fakeSheet('CP sheet', null), ['A', 'B'])).not.toThrow();
});

test('allows a tab whose header matches the expected schema exactly', () => {
  expect(() => assertSafeMirrorTarget_(fakeSheet('CP sheet', ['A', 'B']), ['A', 'B'])).not.toThrow();
});

test('refuses a tab holding a different, unrecognized schema', () => {
  expect(() => assertSafeMirrorTarget_(fakeSheet('TOM', ['Issue Type', 'Issue']), ['A', 'B']))
    .toThrow(/Refusing to mirror/);
});

test('checks only the overlapping width, so a wider real header than expected is still safe', () => {
  expect(() => assertSafeMirrorTarget_(fakeSheet('CP sheet', ['A', 'B', 'Extra']), ['A', 'B'])).not.toThrow();
});
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npx jest tests/MirrorTargetGuard.test.js`
Expected: FAIL — `Cannot find module '../src/MirrorTargetGuard'`.

- [ ] **Step 3: Write `src/MirrorTargetGuard.js`**

```js
// Refuses to let a mirror overwrite a tab holding content it doesn't
// recognize. Generalized over the expected column list rather than a
// hardcoded schema, so any future second mirror in this project could share
// it. This is what stops a typo in CONFIG.CP_MIRROR_SHEET_NAME from
// destroying, say, MASTER SERVICE's hand-curated TOM tab.
function assertSafeMirrorTarget_(sheet, expectedColumns) {
  if (sheet.getLastRow() === 0) {
    return;
  }
  var width = Math.min(sheet.getLastColumn(), expectedColumns.length);
  var header = sheet.getRange(1, 1, 1, width).getValues()[0];
  var matches = expectedColumns.slice(0, width).every(function (col, i) {
    return header[i] === col;
  });
  if (!matches) {
    throw new Error('Refusing to mirror into "' + sheet.getName() + '": it already '
      + 'contains data whose header row does not match the expected schema, so '
      + 'overwriting it would destroy it. Point the mirror at an empty or '
      + 'previously-mirrored tab, or clear that tab deliberately first.');
  }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { assertSafeMirrorTarget_: assertSafeMirrorTarget_ };
}
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `npx jest tests/MirrorTargetGuard.test.js`
Expected: PASS, 4 tests.

- [ ] **Step 5: Commit**

```bash
git add src/MirrorTargetGuard.js tests/MirrorTargetGuard.test.js
git commit -m "feat: add the mirror safety guard" -m "Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 6: Alerts.js — email error alerts

*GAS-only (`MailApp`, `UrlFetchApp`). Not required by Jest — verified for real during Task 12.*

**Files:**
- Create: `src/Alerts.js`

**Interfaces:**
- Consumes: `CONFIG.ALERT_EMAIL`, `CONFIG.SLACK_WEBHOOK_URL` (Config.js, Task 1)
- Produces: `sendErrorAlert_(subject, message)` — `(string, string) -> void`

- [ ] **Step 1: Write `src/Alerts.js`**

```js
function sendErrorAlert_(subject, message) {
  MailApp.sendEmail(CONFIG.ALERT_EMAIL, '[CP Task Automation] ' + subject, message);
  if (CONFIG.SLACK_WEBHOOK_URL) {
    UrlFetchApp.fetch(CONFIG.SLACK_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: '*' + subject + '*\n' + message })
    });
  }
}
```

- [ ] **Step 2: Confirm the pure-module suite is unaffected**

Run: `npx jest`
Expected: same pass count as the end of Task 5 (this file has no tests of its own — it's GAS-only glue, verified for real in Task 12).

- [ ] **Step 3: Commit**

```bash
git add src/Alerts.js
git commit -m "feat: add email error alerts" -m "Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 7: SheetHelpers.js — grid capacity helper

*GAS-only (`Sheet` methods only, no `SpreadsheetApp` calls of its own). Not required by Jest — verified for real during Task 12.*

**Files:**
- Create: `src/SheetHelpers.js`

**Interfaces:**
- Consumes: nothing
- Produces: `cpEnsureGridCapacity_(sheet, requiredCols, requiredRows)` — `(sheet, number, number) -> void`, expands a sheet's grid if it's smaller than required (Apps Script's `setValues` throws if the target range exceeds the sheet's current row/column count)

- [ ] **Step 1: Write `src/SheetHelpers.js`**

```js
function cpEnsureGridCapacity_(sheet, requiredCols, requiredRows) {
  if (sheet.getMaxColumns() < requiredCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredCols - sheet.getMaxColumns());
  }
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }
}
```

- [ ] **Step 2: Confirm the pure-module suite is unaffected**

Run: `npx jest`
Expected: same pass count as Task 6.

- [ ] **Step 3: Commit**

```bash
git add src/SheetHelpers.js
git commit -m "feat: add grid-capacity helper" -m "Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 8: CpConsolidate.js — read, validate, normalize, and write CP Master

*GAS-only (`SpreadsheetApp`, `Logger`). Not required by Jest — verified for real during Task 12.*

**Files:**
- Create: `src/CpConsolidate.js`

**Interfaces:**
- Consumes: `CONFIG` (Config.js), `CP_SOURCE_COLUMNS`/`CP_MASTER_COLUMNS`/`findMissingCpColumns`/`mapCpRow` (CpSchema.js), `normalizeCpRow`/`compareCpRowsByCreatedOnDesc` (CpNormalize.js), `cpEnsureGridCapacity_` (SheetHelpers.js), `sendErrorAlert_` (Alerts.js), `mirrorCpMaster` (CpMirror.js, Task 9) — all bare global references (Apps Script concatenates every file in `src/` into one global scope, so no imports are needed between GAS-only files)
- Produces: `consolidateCpTasks()` — the trigger entry point; `resolveCpSourceTab_`, `readCpSourceTab_`, `ensureCpMasterSheet_`, `writeCpMaster_`, `ensureCpSyncLogSheet_`, `appendCpSyncLogEntry_` — internal helpers

- [ ] **Step 1: Write `src/CpConsolidate.js`**

```js
// Reads "CP Tasks - 2025" and "CP Tasks 2026" from the Channel Partner
// workbook this project does not own, normalizes every row (see
// CpNormalize.js for the four fixes applied), and does a full rebuild of the
// "CP Master" tab — never an upsert, because no column or combination of
// columns in the source data is a reliable unique key (verified during
// design).
//
// Full rebuild means every run replaces CP Master's entire contents. A run
// that reads zero rows across both tabs is treated as a failure, not a valid
// empty state, so a permission problem or a renamed tab can never silently
// blank the tab out.

// Resolves a configured tab name against the workbook's ACTUAL tabs by exact,
// case-sensitive name comparison — deliberately not getSheetByName, whose
// case-insensitive matching has previously caused real data loss on a
// different project's sheet (a typo'd tab name silently resolving to the
// wrong existing tab).
function resolveCpSourceTab_(ss, tabName) {
  var sheet = ss.getSheets().filter(function (s) { return s.getName() === tabName; })[0];
  if (!sheet) {
    throw new Error('CP source tab not found: "' + tabName + '". Check '
      + 'CONFIG.CP_SOURCE_TAB_NAMES against the actual tab names in the '
      + 'source workbook.');
  }
  return sheet;
}

// One source tab -> { rows: [normalized row, ...], tatRepairedCount: N }.
// tatRepairedCount counts rows whose source "TAT in days" was a negative
// number (the diagnosed broken-formula symptom) that normalizeCpRow
// corrected to something else — purely a reporting figure for the Sync Log.
function readCpSourceTab_(sheet, tabName) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) {
    return { rows: [], tatRepairedCount: 0 };
  }
  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0];
  var missing = findMissingCpColumns(headers);
  if (missing.length) {
    throw new Error('Source tab "' + tabName + '" is missing expected column(s): '
      + missing.join(', '));
  }
  var dataRows = values.slice(1).filter(function (rowValues) {
    return rowValues.some(function (v) { return v !== '' && v !== null; });
  });
  var tatRepairedCount = 0;
  var rows = dataRows.map(function (rowValues) {
    var raw = mapCpRow(headers, rowValues);
    var normalized = normalizeCpRow(raw, tabName);
    var sourceTat = Number(raw['TAT in days']);
    if (isFinite(sourceTat) && sourceTat < 0 && normalized['TAT in days'] !== sourceTat) {
      tatRepairedCount += 1;
    }
    return normalized;
  });
  return { rows: rows, tatRepairedCount: tatRepairedCount };
}

function ensureCpMasterSheet_(ss) {
  var sheet = ss.getSheetByName(CONFIG.CP_MASTER_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.CP_MASTER_SHEET_NAME);
  }
  return sheet;
}

// Full replace: overwrite the header and every row in one setValues call,
// then trim any rows left over from a larger previous run.
function writeCpMaster_(sheet, rows) {
  var numCols = CP_MASTER_COLUMNS.length;
  var numRows = rows.length + 1; // +1 for the header row
  cpEnsureGridCapacity_(sheet, numCols, numRows);

  var values = [CP_MASTER_COLUMNS].concat(rows.map(function (row) {
    return CP_MASTER_COLUMNS.map(function (col) {
      return row[col] === undefined ? '' : row[col];
    });
  }));
  sheet.getRange(1, 1, numRows, numCols).setValues(values);
  sheet.getRange(1, 1, 1, numCols).setFontWeight('bold');
  sheet.setFrozenRows(1);

  var lastRow = sheet.getLastRow();
  if (lastRow > numRows) {
    sheet.deleteRows(numRows + 1, lastRow - numRows);
  }
}

function ensureCpSyncLogSheet_(ss) {
  var sheet = ss.getSheetByName(CONFIG.CP_SYNC_LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.CP_SYNC_LOG_SHEET_NAME);
    var header = ['Timestamp', 'Rows Read (2025 tab)', 'Rows Read (2026 tab)', 'Rows Written', 'TAT Values Repaired', 'Status'];
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
  }
  return sheet;
}

function appendCpSyncLogEntry_(ss, entry) {
  var sheet = ensureCpSyncLogSheet_(ss);
  sheet.appendRow([
    entry.timestamp, entry.rowsRead2025, entry.rowsRead2026,
    entry.rowsWritten, entry.tatRepaired, entry.status
  ]);
}

function consolidateCpTasks() {
  var cpWorkbook = SpreadsheetApp.openById(CONFIG.CP_WORKBOOK_SHEET_ID);
  try {
    var sourceSs = SpreadsheetApp.openById(CONFIG.CP_SOURCE_SHEET_ID);
    var rowsByTab = {};
    var allRows = [];
    var totalTatRepaired = 0;

    CONFIG.CP_SOURCE_TAB_NAMES.forEach(function (tabName) {
      var sheet = resolveCpSourceTab_(sourceSs, tabName);
      var result = readCpSourceTab_(sheet, tabName);
      rowsByTab[tabName] = result.rows.length;
      totalTatRepaired += result.tatRepairedCount;
      allRows = allRows.concat(result.rows);
    });

    // CONFIG.CP_SOURCE_TAB_NAMES is defined as [2025 tab, 2026 tab] — the
    // Sync Log's two "Rows Read" columns rely on that order.
    var rowsRead2025 = rowsByTab[CONFIG.CP_SOURCE_TAB_NAMES[0]] || 0;
    var rowsRead2026 = rowsByTab[CONFIG.CP_SOURCE_TAB_NAMES[1]] || 0;

    if (allRows.length === 0) {
      throw new Error('Read 0 rows across all CP source tabs.');
    }

    allRows.sort(compareCpRowsByCreatedOnDesc);

    var target = ensureCpMasterSheet_(cpWorkbook);
    writeCpMaster_(target, allRows);

    appendCpSyncLogEntry_(cpWorkbook, {
      timestamp: new Date(),
      rowsRead2025: rowsRead2025,
      rowsRead2026: rowsRead2026,
      rowsWritten: allRows.length,
      tatRepaired: totalTatRepaired,
      status: 'Success'
    });
    Logger.log('CP consolidation COMPLETE — %s row(s) written to "%s".',
      allRows.length, CONFIG.CP_MASTER_SHEET_NAME);
  } catch (err) {
    var errMessage = (err && err.message) || String(err);
    sendErrorAlert_('CP consolidation failed', errMessage);
    appendCpSyncLogEntry_(cpWorkbook, {
      timestamp: new Date(), rowsRead2025: '', rowsRead2026: '',
      rowsWritten: 0, tatRepaired: 0, status: 'Error: ' + errMessage
    });
    return;
  }

  mirrorCpMaster();
}
```

- [ ] **Step 2: Confirm the pure-module suite is unaffected**

Run: `npx jest`
Expected: same pass count as Task 7.

- [ ] **Step 3: Commit**

```bash
git add src/CpConsolidate.js
git commit -m "feat: add CpConsolidate.js to build CP Master from the two CP Tasks tabs" -m "Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 9: CpMirror.js — mirror CP Master into MASTER SERVICE

*GAS-only (`SpreadsheetApp`, `Logger`). Not required by Jest — verified for real during Task 12.*

**Files:**
- Create: `src/CpMirror.js`

**Interfaces:**
- Consumes: `CONFIG` (Config.js), `CP_MASTER_COLUMNS` (CpSchema.js), `assertSafeMirrorTarget_` (MirrorTargetGuard.js, Task 5), `cpEnsureGridCapacity_` (SheetHelpers.js), `sendErrorAlert_` (Alerts.js) — all bare global references
- Produces: `mirrorCpMaster()` — callable both as the last step of `consolidateCpTasks()` (Task 8) and on its own from the Apps Script editor's function dropdown

- [ ] **Step 1: Write `src/CpMirror.js`**

```js
// One-way mirror of CP Master into the MASTER SERVICE workbook's "CP sheet"
// tab, alongside that workbook's own "Service sheet" and hand-curated "TOM"
// tab — both belonging to the separate ServiceWRK ticket automation and
// never touched by this function.
//
// A few hundred rows comfortably fits one getValues()/setValues() pair, so
// unlike a large mirror this needs no chunking or resume logic.
function mirrorCpMaster() {
  var cpWorkbook = SpreadsheetApp.openById(CONFIG.CP_WORKBOOK_SHEET_ID);
  var source = cpWorkbook.getSheetByName(CONFIG.CP_MASTER_SHEET_NAME);
  if (!source || source.getLastRow() < 2) {
    Logger.log('CP Master has no data rows — nothing to mirror.');
    return;
  }

  try {
    var targetSs = SpreadsheetApp.openById(CONFIG.MASTER_SERVICE_SHEET_ID);
    var target = targetSs.getSheetByName(CONFIG.CP_MIRROR_SHEET_NAME);
    if (!target) {
      target = targetSs.insertSheet(CONFIG.CP_MIRROR_SHEET_NAME);
      Logger.log('Created tab "%s" in "%s".', CONFIG.CP_MIRROR_SHEET_NAME, targetSs.getName());
    }
    assertSafeMirrorTarget_(target, CP_MASTER_COLUMNS);

    var numCols = CP_MASTER_COLUMNS.length;
    var numRows = source.getLastRow(); // includes the header row
    var values = source.getRange(1, 1, numRows, numCols).getValues();

    cpEnsureGridCapacity_(target, numCols, numRows);
    target.getRange(1, 1, numRows, numCols).setValues(values);
    target.getRange(1, 1, 1, numCols).setFontWeight('bold');
    target.setFrozenRows(1);

    var targetLastRow = target.getLastRow();
    if (targetLastRow > numRows) {
      target.deleteRows(numRows + 1, targetLastRow - numRows);
    }
    Logger.log('CP mirror COMPLETE — %s row(s) now in "%s" of "%s".',
      numRows - 1, CONFIG.CP_MIRROR_SHEET_NAME, targetSs.getName());
  } catch (err) {
    var errMessage = (err && err.message) || String(err);
    sendErrorAlert_('CP mirror failed', errMessage);
  }
}
```

- [ ] **Step 2: Confirm the pure-module suite is unaffected**

Run: `npx jest`
Expected: same pass count as Task 8.

- [ ] **Step 3: Commit**

```bash
git add src/CpMirror.js
git commit -m "feat: add CpMirror.js to mirror CP Master into MASTER SERVICE" -m "Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 10: Triggers.js — entry points and schedule installer

*GAS-only (`ScriptApp`). Verified for real during Task 13.*

**Files:**
- Create: `src/Triggers.js`

**Interfaces:**
- Consumes: `CONFIG.CP_TRIGGER_TIMES` (Config.js), `consolidateCpTasks` (CpConsolidate.js, Task 8)
- Produces: `installTriggers()`, `removeAllTriggers_()` — selected from the Apps Script editor's function dropdown in Tasks 12-13

- [ ] **Step 1: Write `src/Triggers.js`**

```js
function removeAllTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    ScriptApp.deleteTrigger(trigger);
  });
}

function installTriggers() {
  removeAllTriggers_();
  // consolidateCpTasks() calls mirrorCpMaster() itself as its last step, so
  // one daily trigger covers the whole pipeline.
  CONFIG.CP_TRIGGER_TIMES.forEach(function (time) {
    ScriptApp.newTrigger('consolidateCpTasks').timeBased().everyDays(1).atHour(time[0]).nearMinute(time[1]).create();
  });
}
```

- [ ] **Step 2: Confirm the pure-module suite is unaffected**

Run: `npx jest`
Expected: same pass count as Task 9.

- [ ] **Step 3: Commit**

```bash
git add src/Triggers.js
git commit -m "feat: add trigger entry points and schedule installer" -m "Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 11: README — document the project

**Files:**
- Create: `README.md`

**Interfaces:**
- Consumes: nothing
- Produces: nothing (documentation only)

- [ ] **Step 1: Write `README.md`**

```markdown
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

## Operating it

Deploy code changes with `npx clasp push --force`.

Functions runnable from the Apps Script editor's dropdown:

| Function | What it does |
|---|---|
| `consolidateCpTasks` | Reads both CP Tasks tabs, rebuilds `CP Master`, then mirrors it |
| `mirrorCpMaster` | Mirrors `CP Master` into MASTER SERVICE's `CP sheet` tab on its own |
| `installTriggers` | (Re)installs the daily trigger; clears existing ones first, safe to re-run |
| `removeAllTriggers_` | Rollback — stops all scheduled runs without touching any data |

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
  SheetHelpers.js      Grid-capacity helper
  CpConsolidate.js    Reads the two CP Tasks tabs, normalizes, rebuilds CP Master
  CpMirror.js         One-way mirror of CP Master into MASTER SERVICE
  Triggers.js         Trigger entry points and installer
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
```

- [ ] **Step 2: Commit**

```bash
git add README.md
git commit -m "docs: add README" -m "Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 12: Deploy and dry-run verification (manual — requires the real Apps Script editor)

*Every step here happens in the Apps Script editor / Google Sheets UI, not in this repo's code. There is no way to execute Google's servers from this environment, so you (the person running this plan) perform these steps and report back what you observed.*

**Files:** none (deployment + verification only)

- [ ] **Step 1: Confirm the script's identity can open the CP source workbook**

The source workbook (`1kF1tDMvK7pmyqFG4ZTHEJ57mab802pw16a1ByBVmvGM`) was shared with
`abi@tricog.com` at one point — confirm it is also accessible to whichever Google
account this project's `clasp` is authenticated as. If not shared with that account,
share it (Viewer is enough) before continuing.

- [ ] **Step 2: Push the code**

Run: `npx clasp push --force`
Expected: lists every `src/*.js` file and `src/appsscript.json` as pushed, no errors.

- [ ] **Step 3: Run the CP consolidation**

In the [Apps Script editor](https://script.google.com/d/1Mq15Zc6qaH-ctxQKmxHh2NvjzKfChurxgaFzQ9_4krWAQlFgiaXYGMBH/edit), select `consolidateCpTasks` from the function dropdown and click Run. Grant the requested OAuth permissions when prompted (Sheets, Mail, external request, script triggers).

Expected: completes without a red error banner.

- [ ] **Step 4: Inspect CP Task Tracker**

Open the [CP Task Tracker](https://docs.google.com/spreadsheets/d/1dWSDJUjafdnAWsbiXSwo9FlZXC_Qb3QqP8Psek6tPYI/edit) workbook:

- A **CP Master** tab should now exist with a bold, frozen header row of 19 columns ending in `Source Tab`, `Created Year`, and one data row per CP visit across both source tabs, sorted by `Created On` descending.
- Spot-check the row whose `Zoho` is `#125704` (the row analyzed during design): `TAT in days` should be blank, not `-46013`; `Ticket Status` should read exactly `Cancelled` with no trailing space.
- A **Sync Log** tab should have one row: `Rows Read (2025 tab)` and `Rows Read (2026 tab)` both non-zero, `Rows Written` equal to their sum, `TAT Values Repaired` non-zero, `Status` = `Success`.

If anything looks wrong, note exactly what (which column, which row) — that's a debugging task, not a plan step to guess at here.

- [ ] **Step 5: Inspect the MASTER SERVICE mirror**

Open the [MASTER SERVICE](https://docs.google.com/spreadsheets/d/16Q2q9R6GPBOBYVmvImRTZRp8g1kW-G6fio26XDJiULo/edit) workbook:

- A **CP sheet** tab should now exist, with the same header and row count as `CP Master`.
- The existing `Service sheet` and `TOM` tabs (belonging to the separate ServiceWRK ticket automation) should be completely unaffected — same row/column counts as before Step 3.

- [ ] **Step 6: Re-run to confirm the full-rebuild path is safe to repeat**

Run `consolidateCpTasks` a second time from the function dropdown.

Expected: `CP Master` ends up with the exact same row count as after Step 3 (not doubled), a second `Sync Log` row is appended (not overwritten), and `CP sheet` still matches `CP Master`.

- [ ] **Step 7: Record the outcome**

No commit needed (nothing in the repo changes) — confirm all the above checks passed before moving to Task 13.

---

### Task 13: Enable the schedule

**Files:** none

- [ ] **Step 1: Install the trigger**

In the Apps Script editor, select `installTriggers` and click Run.

- [ ] **Step 2: Verify**

Open the Apps Script project's Triggers page (clock icon in the left sidebar). Expected: 1 time-driven trigger, running `consolidateCpTasks` daily at 8:30.

- [ ] **Step 3: Note the go-live date and confirm the rollback path**

Confirm the trigger can be deleted individually from the Triggers page, or removed by
running `removeAllTriggers_`. Deleting it stops future runs without touching any data
already written to `CP Master` or its mirror.

- [ ] **Step 4: Watch the first few days**

Watch the `Sync Log` tab in CP Task Tracker and your email for anything the dry run in
Task 12 didn't cover — in particular, a real row count much larger than the 145-per-tab
sample seen during design (that sample was reader-capped, not the true size, so a bigger
real number is expected and not itself a problem).
