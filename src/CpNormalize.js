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
