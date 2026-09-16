function sendErrorAlert_(subject, message) {
  MailApp.sendEmail(CONFIG.ALERT_EMAIL, '[CP Task Automation] ' + subject, message);
  if (CONFIG.SLACK_WEBHOOK_URL) {
    UrlFetchApp.fetch(CONFIG.SLACK_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: '*' + subject + '*\n' + message })
    });
  }
}
