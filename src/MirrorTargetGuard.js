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
