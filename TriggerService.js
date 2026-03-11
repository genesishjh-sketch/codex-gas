/** 트리거 설치/삭제/재설치 유틸 */

function runInteriorDbSyncByTrigger() {
  try {
    runInteriorDbSync();
  } catch (err) {
    console.error('자동 동기화 실패: ' + (err && err.message ? err.message : err));
    throw err;
  }
}

function enableInteriorSyncOnOpen() {
  PropertiesService.getDocumentProperties().setProperty(INTERIOR_SYNC_KEYS.AUTO_SYNC_ON_OPEN, 'true');
  alertIfPossible_(getUiIfAvailable_(), '열 때 자동 동기화 플래그를 켰습니다.');
}

function disableInteriorSyncOnOpen() {
  PropertiesService.getDocumentProperties().setProperty(INTERIOR_SYNC_KEYS.AUTO_SYNC_ON_OPEN, 'false');
  alertIfPossible_(getUiIfAvailable_(), '열 때 자동 동기화를 껐습니다.');
}

function runInteriorSyncOnOpenIfEnabled_() {
  var enabled = PropertiesService.getDocumentProperties().getProperty(INTERIOR_SYNC_KEYS.AUTO_SYNC_ON_OPEN) === 'true';
  if (!enabled) return;
  console.log('열 때 자동 동기화는 비활성 경로입니다. 시간 기반 트리거를 사용하세요.');
}

function installDailyInteriorSyncTrigger6am() {
  removeDailyInteriorSyncTriggers();
  ScriptApp.newTrigger(INTERIOR_SYNC_TRIGGER_HANDLER)
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  alertIfPossible_(getUiIfAvailable_(), '매일 오전 6시 자동 동기화 트리거를 설치했습니다.');
}

function installDailyInteriorSyncTriggerBySettings() {
  var settings = getInteriorMasterSettings_();
  var parsed = parseKstTime_(settings.DAILY_SYNC_TIME_KST || INTERIOR_MASTER_SETTINGS.DEFAULTS.DAILY_SYNC_TIME_KST);

  removeDailyInteriorSyncTriggers();

  var builder = ScriptApp.newTrigger(INTERIOR_SYNC_TRIGGER_HANDLER)
    .timeBased()
    .everyDays(1)
    .atHour(parsed.hour);

  if (typeof builder.nearMinute === 'function') {
    builder = builder.nearMinute(parsed.minute);
  }
  builder.create();

  alertIfPossible_(
    getUiIfAvailable_(),
    '설정 기준 자동 동기화 트리거를 설치했습니다.\n'
    + '- 실행시각(KST): ' + parsed.text + '\n'
    + '- 동기화 범위: ' + getInteriorSyncScopeMode_() + '\n'
    + '- 아카이브 일수: ' + getInteriorArchiveAfterDays_() + '일'
  );
}

function installRealtimeInteriorSyncTrigger() {
  removeRealtimeInteriorSyncTriggers();
  ScriptApp.newTrigger(INTERIOR_REALTIME_SYNC_TRIGGER_HANDLER)
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  alertIfPossible_(getUiIfAvailable_(), '실시간 동기화 트리거를 설치했습니다.\n(변경된 행이 속한 프로젝트만 동기화)');
}

function removeRealtimeInteriorSyncTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === INTERIOR_REALTIME_SYNC_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function removeDailyInteriorSyncTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === INTERIOR_SYNC_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function reinstallDailyInteriorSyncTriggerBySettings() {
  removeDailyInteriorSyncTriggers();
  installDailyInteriorSyncTriggerBySettings();
}
