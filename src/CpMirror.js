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
