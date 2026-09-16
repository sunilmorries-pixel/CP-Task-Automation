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
