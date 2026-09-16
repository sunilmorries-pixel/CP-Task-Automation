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
  try {
    var cpWorkbook = SpreadsheetApp.openById(CONFIG.CP_WORKBOOK_SHEET_ID);
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
    if (cpWorkbook) {
      appendCpSyncLogEntry_(cpWorkbook, {
        timestamp: new Date(), rowsRead2025: '', rowsRead2026: '',
        rowsWritten: 0, tatRepaired: 0, status: 'Error: ' + errMessage
      });
    }
    return;
  }

  mirrorCpMaster();
}
