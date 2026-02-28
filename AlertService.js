/** 텔레그램 브리핑/지연 알림 */

function sendTelegramDailyBriefing() {
  var token = getTelegramBotToken_();
  var chatId = getTelegramChatId_();
  if (!token || !chatId) {
    throw new Error('텔레그램 토큰 또는 chat_id 설정이 없습니다.');
  }

  var stats = refreshInteriorDashboard90d();
  var text = '[인테리어 일일 브리핑]\n'
    + '- 최근90일 작업: ' + stats.total + '\n'
    + '- 최근90일 완료: ' + stats.done + '\n'
    + '- 완료율: ' + stats.completionRate + '%\n'
    + '- 지연: ' + stats.delayed;

  return postTelegramMessage_(token, chatId, text);
}

function sendTelegramDelayAlerts() {
  var token = getTelegramBotToken_();
  var chatId = getTelegramChatId_();
  if (!token || !chatId) {
    throw new Error('텔레그램 토큰 또는 chat_id 설정이 없습니다.');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var milestonesSheet = getOrCreateTargetSheet_(ss, INTERIOR_SYNC_CONFIG.TARGET_MILESTONES_ALIASES, INTERIOR_SYNC_CONFIG.TARGET_MILESTONES, INTERIOR_SYNC_CONFIG.TARGET_HEADERS.milestones);
  var lastRow = milestonesSheet.getLastRow();
  if (lastRow < 2) return { sent: false, delayedCount: 0 };

  var rows = milestonesSheet.getRange(2, 1, lastRow - 1, milestonesSheet.getLastColumn()).getValues();
  var now = new Date();
  var delayed = rows.filter(function(row) {
    var plan = row[3];
    var done = row[4];
    var planDate = plan instanceof Date ? plan : (plan ? new Date(plan) : null);
    var doneDate = done instanceof Date ? done : (done ? new Date(done) : null);
    return planDate && !isNaN(planDate.getTime()) && !doneDate && planDate < now;
  });

  if (delayed.length === 0) return { sent: false, delayedCount: 0 };

  var lines = delayed.slice(0, 20).map(function(row) {
    return '- [' + row[0] + '] ' + row[2] + ' / 담당: ' + (row[5] || '미지정');
  });

  var text = '[지연 알림] 총 ' + delayed.length + '건\n' + lines.join('\n');
  postTelegramMessage_(token, chatId, text);
  return { sent: true, delayedCount: delayed.length };
}

function postTelegramMessage_(token, chatId, text) {
  var url = 'https://api.telegram.org/bot' + encodeURIComponent(token) + '/sendMessage';
  var payload = {
    chat_id: chatId,
    text: text
  };
  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  });
  return { code: response.getResponseCode(), body: response.getContentText() };
}

function getTelegramBotToken_() {
  return (PropertiesService.getScriptProperties().getProperty('INTERIOR_TELEGRAM_BOT_TOKEN') || '').toString().trim();
}

function getTelegramChatId_() {
  return (PropertiesService.getScriptProperties().getProperty('INTERIOR_TELEGRAM_CHAT_ID') || '').toString().trim();
}
