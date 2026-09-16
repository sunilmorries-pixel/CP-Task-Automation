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
