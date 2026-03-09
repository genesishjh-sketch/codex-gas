/** Todoist REST API 통신 모듈 */
function todoistCreateTask_(payload) {
  return todoistRequest_('/tasks', 'post', payload);
}

function todoistUpdateTask_(taskId, payload) {
  if (!taskId) throw new Error('업데이트 대상 taskId가 없습니다.');
  return todoistRequest_('/tasks/' + encodeURIComponent(taskId), 'post', payload);
}

function todoistRequest_(path, method, payload) {
  var tokenInfo = getTodoistApiToken_();
  if (!tokenInfo.token) {
    throw new Error('Todoist API 토큰이 없습니다. settings의 todoist_api_token 또는 Script Properties의 TODOIST_API_TOKEN을 설정하세요.');
  }

  var url = TODOIST_SYNC.TODOIST_API_BASE_URL + path;
  var options = {
    method: method,
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + tokenInfo.token },
    payload: JSON.stringify(payload || {}),
    muteHttpExceptions: true
  };

  Logger.log('[Todoist] %s %s payload=%s', method.toUpperCase(), url, JSON.stringify(payload || {}));
  var res = UrlFetchApp.fetch(url, options);
  var code = res.getResponseCode();
  var body = res.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error('Todoist API 오류(' + code + '): ' + body);
  }

  return body ? JSON.parse(body) : {};
}
