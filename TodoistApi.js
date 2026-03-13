/** Todoist REST API 통신 모듈 */
var TODOIST_PROJECT_COLLABORATORS_CACHE_ = {};

function todoistCreateTask_(payload) {
  return todoistRequest_('/tasks', 'post', payload);
}

function todoistUpdateTask_(taskId, payload) {
  if (!taskId) throw new Error('업데이트 대상 taskId가 없습니다.');
  return todoistRequest_('/tasks/' + encodeURIComponent(taskId), 'post', payload);
}

function todoistGetTask_(taskId) {
  if (!taskId) throw new Error('조회 대상 taskId가 없습니다.');
  return todoistRequest_('/tasks/' + encodeURIComponent(taskId), 'get');
}

function todoistCloseTask_(taskId) {
  if (!taskId) throw new Error('완료 처리 대상 taskId가 없습니다.');
  return todoistRequest_('/tasks/' + encodeURIComponent(taskId) + '/close', 'post', {});
}

function todoistReopenTask_(taskId) {
  if (!taskId) throw new Error('재오픈 대상 taskId가 없습니다.');
  return todoistRequest_('/tasks/' + encodeURIComponent(taskId) + '/reopen', 'post', {});
}

function todoistFindActiveTaskByTaskUid_(projectId, taskUid) {
  var normalizedProjectId = (projectId || '').toString().trim();
  var normalizedUid = (taskUid || '').toString().trim();
  if (!normalizedProjectId || !normalizedUid) return null;

  var marker = 'meta_task_uid:' + normalizedUid;
  var tasks = todoistRequest_('/tasks?project_id=' + encodeURIComponent(normalizedProjectId), 'get');
  if (!tasks || !tasks.length) return null;

  for (var i = 0; i < tasks.length; i++) {
    var task = tasks[i] || {};
    var description = (task.description || '').toString();
    if (description.indexOf(marker) >= 0) return task;
  }

  return null;
}

function todoistGetCompletedTaskByTaskId_(taskId) {
  var normalizedTaskId = (taskId || '').toString().trim();
  if (!normalizedTaskId) return null;

  var cursor = '';
  while (true) {
    var path = '/tasks/completed/by_completion_date?limit=200';
    if (cursor) {
      path += '&cursor=' + encodeURIComponent(cursor);
    }

    var response = todoistRequest_(path, 'get');
    var items = response.items || [];
    for (var i = 0; i < items.length; i++) {
      var item = items[i] || {};
      if ((item.task_id || '').toString().trim() === normalizedTaskId) return item;
    }

    cursor = response.next_cursor || '';
    if (!cursor) break;
  }

  return null;
}

function todoistRequest_(path, method, payload) {
  var tokenInfo = getTodoistApiToken_();
  if (!tokenInfo.token) {
    throw new Error('Todoist API 토큰이 없습니다. settings의 todoist_api_token 또는 Script Properties의 TODOIST_API_TOKEN을 설정하세요.');
  }

  var url = TODOIST_SYNC.TODOIST_API_BASE_URL + path;
  var options = {
    method: method,
    headers: { Authorization: 'Bearer ' + tokenInfo.token },
    muteHttpExceptions: true
  };

  if (method.toLowerCase() !== 'get') {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(payload || {});
  }

  Logger.log('[Todoist] %s %s payload=%s', method.toUpperCase(), url, JSON.stringify(payload || {}));
  var res = UrlFetchApp.fetch(url, options);
  var code = res.getResponseCode();
  var body = res.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error('Todoist API 오류(' + code + '): ' + body);
  }

  return body ? JSON.parse(body) : {};
}

function todoistFindCollaboratorIdByEmail_(projectId, email) {
  var normalizedProjectId = (projectId || '').toString().trim();
  var normalizedEmail = (email || '').toString().trim().toLowerCase();
  if (!normalizedProjectId || !normalizedEmail) return '';

  var collaborators = getTodoistProjectCollaborators_(normalizedProjectId);
  for (var i = 0; i < collaborators.length; i++) {
    var collaborator = collaborators[i] || {};
    var collaboratorEmail = (collaborator.email || '').toString().trim().toLowerCase();
    if (collaboratorEmail === normalizedEmail) {
      return (collaborator.id || '').toString().trim();
    }
  }

  return '';
}

function getTodoistProjectCollaborators_(projectId) {
  if (TODOIST_PROJECT_COLLABORATORS_CACHE_.hasOwnProperty(projectId)) {
    return TODOIST_PROJECT_COLLABORATORS_CACHE_[projectId];
  }

  var collaborators = [];
  var cursor = '';

  while (true) {
    var query = '?limit=200';
    if (cursor) {
      query += '&cursor=' + encodeURIComponent(cursor);
    }

    var response = todoistRequest_('/projects/' + encodeURIComponent(projectId) + '/collaborators' + query, 'get');
    var results = response.results || [];
    collaborators = collaborators.concat(results);

    cursor = response.next_cursor || '';
    if (!cursor) break;
  }

  TODOIST_PROJECT_COLLABORATORS_CACHE_[projectId] = collaborators;
  return collaborators;
}
