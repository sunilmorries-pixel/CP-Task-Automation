function cpEnsureGridCapacity_(sheet, requiredCols, requiredRows) {
  if (sheet.getMaxColumns() < requiredCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredCols - sheet.getMaxColumns());
  }
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }
}

// Resolves a tab by exact, case-sensitive name — never getSheetByName, whose
// case-insensitive matching has previously caused real data loss elsewhere (a
// typo'd tab name silently resolving to the wrong existing tab). Returns null
// rather than throwing when no exact match exists; callers decide whether
// "missing" means "create it" or "fail the run".
function resolveSheetByExactName_(ss, name) {
  return ss.getSheets().filter(function (s) { return s.getName() === name; })[0] || null;
}
