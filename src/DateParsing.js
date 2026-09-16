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
