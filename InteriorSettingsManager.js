/**
 * 인테리어 동기화 설정 전용 모듈
 * - Script Properties 기반으로 설정 조회
 */

var INTERIOR_MASTER_SETTINGS = {
  DEFAULTS: {
    DAILY_SYNC_TIME_KST: '08:30',
    SYNC_SCOPE_MODE: '지연+7일예정',
    ARCHIVE_AFTER_DAYS: '30'
  },
  SCOPE_OPTIONS: {
    '지연만': true,
    '7일예정만': true,
    '지연+7일예정': true,
    '전체': true
  }
};

function getInteriorMasterSettings_() {
  var props = PropertiesService.getScriptProperties();
  var defaults = INTERIOR_MASTER_SETTINGS.DEFAULTS;

  return {
    DAILY_SYNC_TIME_KST: (props.getProperty('DAILY_SYNC_TIME_KST') || defaults.DAILY_SYNC_TIME_KST).toString().trim(),
    SYNC_SCOPE_MODE: (props.getProperty('SYNC_SCOPE_MODE') || defaults.SYNC_SCOPE_MODE).toString().trim(),
    ARCHIVE_AFTER_DAYS: (props.getProperty('ARCHIVE_AFTER_DAYS') || defaults.ARCHIVE_AFTER_DAYS).toString().trim()
  };
}

function getInteriorSyncScopeMode_() {
  var mode = getInteriorMasterSettings_().SYNC_SCOPE_MODE;
  return INTERIOR_MASTER_SETTINGS.SCOPE_OPTIONS[mode] ? mode : INTERIOR_MASTER_SETTINGS.DEFAULTS.SYNC_SCOPE_MODE;
}

function getInteriorArchiveAfterDays_() {
  var raw = getInteriorMasterSettings_().ARCHIVE_AFTER_DAYS;
  var n = parseInt(raw, 10);
  return (n >= 1) ? n : parseInt(INTERIOR_MASTER_SETTINGS.DEFAULTS.ARCHIVE_AFTER_DAYS, 10);
}

function parseKstTime_(value) {
  var text = (value || '').toString().trim();
  var match = text.match(/^([01]\d|2[0-3]):([0-5]\d)$/);
  if (!match) {
    var fallback = INTERIOR_MASTER_SETTINGS.DEFAULTS.DAILY_SYNC_TIME_KST;
    match = fallback.match(/^([01]\d|2[0-3]):([0-5]\d)$/);
  }
  return {
    hour: parseInt(match[1], 10),
    minute: parseInt(match[2], 10),
    text: ('0' + match[1]).slice(-2) + ':' + ('0' + match[2]).slice(-2)
  };
}
