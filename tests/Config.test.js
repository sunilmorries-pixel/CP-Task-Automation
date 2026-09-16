const { CONFIG } = require('../src/Config');

test('CP source workbook and tabs point at the approved Channel Partner tracker', () => {
  expect(CONFIG.CP_SOURCE_SHEET_ID).toBe('1kF1tDMvK7pmyqFG4ZTHEJ57mab802pw16a1ByBVmvGM');
  expect(CONFIG.CP_SOURCE_TAB_NAMES).toEqual(['CP Tasks - 2025', 'CP Tasks 2026']);
});

test('CP workbook and mirror target match the approved design spec', () => {
  expect(CONFIG.CP_WORKBOOK_SHEET_ID).toBe('1dWSDJUjafdnAWsbiXSwo9FlZXC_Qb3QqP8Psek6tPYI');
  expect(CONFIG.CP_MASTER_SHEET_NAME).toBe('CP Master');
  expect(CONFIG.CP_SYNC_LOG_SHEET_NAME).toBe('Sync Log');
  expect(CONFIG.MASTER_SERVICE_SHEET_ID).toBe('16Q2q9R6GPBOBYVmvImRTZRp8g1kW-G6fio26XDJiULo');
  expect(CONFIG.CP_MIRROR_SHEET_NAME).toBe('CP sheet');
  expect(CONFIG.CP_TRIGGER_TIMES).toHaveLength(1);
});

test('the CP workbook, source workbook, and mirror target are three distinct resources', () => {
  expect(CONFIG.CP_WORKBOOK_SHEET_ID).not.toBe(CONFIG.CP_SOURCE_SHEET_ID);
  expect(CONFIG.CP_WORKBOOK_SHEET_ID).not.toBe(CONFIG.MASTER_SERVICE_SHEET_ID);
  expect(CONFIG.CP_SOURCE_SHEET_ID).not.toBe(CONFIG.MASTER_SERVICE_SHEET_ID);
});

test('the CP mirror tab name cannot collide with CP Master or Sync Log, case-insensitively', () => {
  const mirror = CONFIG.CP_MIRROR_SHEET_NAME.toLowerCase();
  expect(mirror).not.toBe(CONFIG.CP_MASTER_SHEET_NAME.toLowerCase());
  expect(mirror).not.toBe(CONFIG.CP_SYNC_LOG_SHEET_NAME.toLowerCase());
});
