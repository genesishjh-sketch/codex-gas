/** Todoist REST API 통신 모듈 */
function todoistCreateTask_(payload) {
  return todoistRequest_('/tasks', 'post', payload);
}

function todoistUpdateTask_(taskId, payload) {
  if (!taskId) throw new Error('업데이트 대상 taskId가 없습니다.');
  return todoistRequest_('/tasks/' + encodeURIComponent(taskId), 'post', payload);
}

function todoistRequest_(path, method, payload) {
  var token = PropertiesService.getScriptProperties().getProperty(TODOIST_SYNC.PROPERTY_API_TOKEN);
  if (!token) {
    throw new Error('Script Properties에 TODOIST_API_TOKEN이 없습니다.');
  }

  var url = TODOIST_SYNC.TODOIST_API_BASE_URL + path;
  var options = {
    method: method,
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
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
