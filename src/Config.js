var CONFIG = {
  // Source: read-only, owned outside this project.
  CP_SOURCE_SHEET_ID: '1kF1tDMvK7pmyqFG4ZTHEJ57mab802pw16a1ByBVmvGM',
  CP_SOURCE_TAB_NAMES: ['CP Tasks - 2025', 'CP Tasks 2026'],

  // Build target: this project's own workbook.
  CP_WORKBOOK_SHEET_ID: '1dWSDJUjafdnAWsbiXSwo9FlZXC_Qb3QqP8Psek6tPYI',
  CP_MASTER_SHEET_NAME: 'CP Master',
  CP_SYNC_LOG_SHEET_NAME: 'Sync Log',

  // Mirror target: MASTER SERVICE, belonging to the separate ServiceWRK ticket
  // automation. Only CP_MIRROR_SHEET_NAME is ever written; every other tab in
  // that workbook belongs to that other project and must never be touched.
  MASTER_SERVICE_SHEET_ID: '16Q2q9R6GPBOBYVmvImRTZRp8g1kW-G6fio26XDJiULo',
  CP_MIRROR_SHEET_NAME: 'CP sheet',

  CP_TRIGGER_TIMES: [[8, 30]],
  ALERT_EMAIL: 'sunil.morries@tricog.com',
  SLACK_WEBHOOK_URL: ''
};

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { CONFIG: CONFIG };
}
