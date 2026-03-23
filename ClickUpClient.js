/** ClickUpClient.gs
 * ClickUp API 호출 모듈.
 */

function createClickUpClient_(settings) {
  var token = stringValue_(settings.auth.CLICKUP_API_TOKEN);
  var listId = stringValue_(settings.target.CLICKUP_LIST_ID);

  function request_(method, path, payload) {
    var url = 'https://api.clickup.com/api/v2' + path;
    var options = {
      method: method,
      contentType: 'application/json',
      muteHttpExceptions: true,
      headers: {
        Authorization: token
      }
    };
    if (payload) options.payload = JSON.stringify(payload);

    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    var body = response.getContentText();
    var json = {};
    try { json = body ? JSON.parse(body) : {}; } catch (e) {}

    if (code < 200 || code >= 300) {
      throw new Error('ClickUp API 실패 (' + code + '): ' + body);
    }
    return json;
  }

  return {
    createParentTask: function(taskName, description, status, dueDateMs) {
      var payload = {
        name: taskName,
        description: description
      };
      if (status) payload.status = status;
      if (dueDateMs) payload.due_date = dueDateMs;

      return request_('post', '/list/' + listId + '/task', payload);
    },

    createSubtask: function(parentTaskId, taskName, description, status, dueDateMs) {
      var payload = {
        name: taskName,
        description: description,
        parent: parentTaskId
      };
      if (status) payload.status = status;
      if (dueDateMs) payload.due_date = dueDateMs;

      return request_('post', '/list/' + listId + '/task', payload);
    }
  };
}
