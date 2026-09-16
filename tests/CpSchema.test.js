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
