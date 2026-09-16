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
