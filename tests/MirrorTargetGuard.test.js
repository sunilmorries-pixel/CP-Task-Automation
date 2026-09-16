const { assertSafeMirrorTarget_ } = require('../src/MirrorTargetGuard');

function fakeSheet(name, header) {
  return {
    getName: function () { return name; },
    getLastRow: function () { return header ? 5 : 0; },
    getLastColumn: function () { return header ? header.length : 0; },
    getRange: function () {
      return { getValues: function () { return [header]; } };
    }
  };
}

test('allows a completely empty tab', () => {
  expect(() => assertSafeMirrorTarget_(fakeSheet('CP sheet', null), ['A', 'B'])).not.toThrow();
});

test('allows a tab whose header matches the expected schema exactly', () => {
  expect(() => assertSafeMirrorTarget_(fakeSheet('CP sheet', ['A', 'B']), ['A', 'B'])).not.toThrow();
});

test('refuses a tab holding a different, unrecognized schema', () => {
  expect(() => assertSafeMirrorTarget_(fakeSheet('TOM', ['Issue Type', 'Issue']), ['A', 'B']))
    .toThrow(/Refusing to mirror/);
});

test('checks only the overlapping width, so a wider real header than expected is still safe', () => {
  expect(() => assertSafeMirrorTarget_(fakeSheet('CP sheet', ['A', 'B', 'Extra']), ['A', 'B'])).not.toThrow();
});
