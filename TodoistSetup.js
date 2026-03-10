/** Todoist 연동 설치/초기화 모듈 */
function setupTodoistMilestonesSync() {
  ensureTodoistSettingsLayout_();
  var settings = getTodoistSyncSettings_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var target = ss.getSheetByName(settings.sync_target_sheet || TODOIST_SYNC.DEFAULT_TARGET_SHEET);
  if (!target) throw new Error('동기화 대상 시트가 없습니다: ' + settings.sync_target_sheet);

  ensureMilestonesSyncColumns_(target);
  installTodoistEditTrigger_();

  var tokenInfo = getTodoistApiToken_();
  var message = [
    'Todoist 동기화 설치 완료',
    '- 대상 시트: ' + target.getName(),
    '- 실시간 트리거: 설치됨',
    '- API 토큰: ' + (tokenInfo.token ? '설정됨 (' + (tokenInfo.source === 'settings' ? 'settings 시트' : 'Script Properties') + ')' : '미설정 (settings 시트 todoist_api_token 또는 Script Properties TODOIST_API_TOKEN 설정 필요)')
  ].join('\n');

  alertIfPossible_(getUiIfAvailable_(), message);
}

function installTodoistEditTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  var exists = triggers.some(function(trigger) {
    return trigger.getHandlerFunction() === TODOIST_SYNC.INSTALLABLE_EDIT_TRIGGER_HANDLER;
  });
  if (!exists) {
    ScriptApp.newTrigger(TODOIST_SYNC.INSTALLABLE_EDIT_TRIGGER_HANDLER)
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();
  }
}

function installDailyTodoistSyncTrigger9am() {
  removeDailyTodoistSyncTriggers();
  ScriptApp.newTrigger(TODOIST_SYNC.DAILY_TRIGGER_HANDLER)
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  alertIfPossible_(getUiIfAvailable_(), '매일 오전 자동 Todoist 동기화 트리거(09시)를 설치했습니다.');
}

function removeDailyTodoistSyncTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === TODOIST_SYNC.DAILY_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}
