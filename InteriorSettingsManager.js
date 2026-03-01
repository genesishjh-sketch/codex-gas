/**
 * 인테리어 동기화 설정 전용 모듈
 * - 마스터설정 시트 생성/검증
 * - 설정값 조회/정규화
 * - 설정 기반 트리거 설치
 */

var INTERIOR_MASTER_SETTINGS = {
  SHEET_NAME: '마스터설정',
  HEADER: ['항목키', '설정값', '설명'],
  DEFAULTS: {
    TODOIST_API_TOKEN: '',
    TODOIST_PROJECT_ID: '',
    TODOIST_TEST_PROJECT_ID: '',
    DAILY_SYNC_TIME_KST: '08:30',
    SYNC_SCOPE_MODE: '지연+7일예정',
    ARCHIVE_AFTER_DAYS: '30'
  },
  SCOPE_OPTIONS_TEXT: '허용값: 지연만 / 7일예정만 / 지연+7일예정 / 전체',
  SCOPE_OPTIONS: {
    '지연만': true,
    '7일예정만': true,
    '지연+7일예정': true,
    '전체': true
  }
};

function setupInteriorMasterSettingsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(INTERIOR_MASTER_SETTINGS.SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(INTERIOR_MASTER_SETTINGS.SHEET_NAME);

  sheet.getRange(1, 1, 1, INTERIOR_MASTER_SETTINGS.HEADER.length).setValues([INTERIOR_MASTER_SETTINGS.HEADER]);

  var rows = [
    ['TODOIST_API_TOKEN', INTERIOR_MASTER_SETTINGS.DEFAULTS.TODOIST_API_TOKEN, 'Todoist 개인 API 토큰. 코드가 아니라 이 셀에서 읽습니다.'],
    ['TODOIST_PROJECT_ID', INTERIOR_MASTER_SETTINGS.DEFAULTS.TODOIST_PROJECT_ID, 'Todoist 기본 프로젝트 ID(실운영).'],
    ['TODOIST_TEST_PROJECT_ID', INTERIOR_MASTER_SETTINGS.DEFAULTS.TODOIST_TEST_PROJECT_ID, 'Todoist 테스트 프로젝트 ID(테스트 태스크 생성용).'],
    ['DAILY_SYNC_TIME_KST', INTERIOR_MASTER_SETTINGS.DEFAULTS.DAILY_SYNC_TIME_KST, '매일 동기화 시각(KST, 24시간 HH:mm). 예: 08:30 / 23:30'],
    ['SYNC_SCOPE_MODE', INTERIOR_MASTER_SETTINGS.DEFAULTS.SYNC_SCOPE_MODE, INTERIOR_MASTER_SETTINGS.SCOPE_OPTIONS_TEXT],
    ['ARCHIVE_AFTER_DAYS', INTERIOR_MASTER_SETTINGS.DEFAULTS.ARCHIVE_AFTER_DAYS, '완료 후 보관으로 이동할 기준 일수(숫자). 예: 30']
  ];

  var existingLastRow = sheet.getLastRow();
  var existingMap = {};
  if (existingLastRow >= 2) {
    var existing = sheet.getRange(2, 1, existingLastRow - 1, 3).getValues();
    existing.forEach(function(row) {
      var key = (row[0] || '').toString().trim();
      if (key) existingMap[key] = row;
    });
  }

  var finalRows = rows.map(function(row) {
    var key = row[0];
    if (!existingMap[key]) return row;
    var currentValue = (existingMap[key][1] || '').toString().trim();
    return [key, currentValue || row[1], row[2]];
  });

  sheet.getRange(2, 1, finalRows.length, 3).setValues(finalRows);

  if (sheet.getLastRow() > finalRows.length + 1) {
    sheet.getRange(finalRows.length + 2, 1, sheet.getLastRow() - (finalRows.length + 1), Math.max(sheet.getLastColumn(), 3)).clearContent();
  }

  var tokenCell = sheet.getRange(2, 2, 1, 1);
  var tokenRule = SpreadsheetApp.newDataValidation().setAllowInvalid(true).build();
  tokenCell.setDataValidation(tokenRule);

  var idRule = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied('=LEN(B3)>0')
    .setAllowInvalid(false)
    .setHelpText('Todoist 프로젝트 ID를 입력하세요.')
    .build();
  sheet.getRange(3, 2).setDataValidation(idRule);
  var testIdRule = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied('=LEN(B4)>0')
    .setAllowInvalid(false)
    .setHelpText('Todoist 프로젝트 ID를 입력하세요.')
    .build();
  sheet.getRange(4, 2).setDataValidation(testIdRule);

  var timeRule = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied('=REGEXMATCH(B5,"^([01]\\\\d|2[0-3]):([0-5]\\\\d)$")')
    .setAllowInvalid(false)
    .setHelpText('24시간 HH:mm 형식으로 입력하세요. 예: 08:30')
    .build();
  sheet.getRange(5, 2).setDataValidation(timeRule);

  var scopeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.keys(INTERIOR_MASTER_SETTINGS.SCOPE_OPTIONS), true)
    .setAllowInvalid(false)
    .setHelpText(INTERIOR_MASTER_SETTINGS.SCOPE_OPTIONS_TEXT)
    .build();
  sheet.getRange(6, 2).setDataValidation(scopeRule);

  var dayRule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThanOrEqualTo(1)
    .setAllowInvalid(false)
    .setHelpText('1 이상의 숫자를 입력하세요. 예: 30')
    .build();
  sheet.getRange(7, 2).setDataValidation(dayRule);

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 3);
  alertIfPossible_(
    getUiIfAvailable_(),
    '마스터설정 탭을 준비했습니다.\n- Todoist 토큰\n- Todoist 실운영/테스트 프로젝트 ID\n- 동기화 시간(KST)\n- 동기화 범위\n- 아카이브 일수\n를 입력하세요.'
  );
}

function getInteriorMasterSettings_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(INTERIOR_MASTER_SETTINGS.SHEET_NAME);
  if (!sheet) return INTERIOR_MASTER_SETTINGS.DEFAULTS;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return INTERIOR_MASTER_SETTINGS.DEFAULTS;

  var values = sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
  var settings = {};
  values.forEach(function(row) {
    var key = (row[0] || '').toString().trim();
    if (!key) return;
    settings[key] = (row[1] || '').toString().trim();
  });

  return {
    TODOIST_API_TOKEN: settings.TODOIST_API_TOKEN || INTERIOR_MASTER_SETTINGS.DEFAULTS.TODOIST_API_TOKEN,
    TODOIST_PROJECT_ID: settings.TODOIST_PROJECT_ID || INTERIOR_MASTER_SETTINGS.DEFAULTS.TODOIST_PROJECT_ID,
    TODOIST_TEST_PROJECT_ID: settings.TODOIST_TEST_PROJECT_ID || INTERIOR_MASTER_SETTINGS.DEFAULTS.TODOIST_TEST_PROJECT_ID,
    DAILY_SYNC_TIME_KST: settings.DAILY_SYNC_TIME_KST || INTERIOR_MASTER_SETTINGS.DEFAULTS.DAILY_SYNC_TIME_KST,
    SYNC_SCOPE_MODE: settings.SYNC_SCOPE_MODE || INTERIOR_MASTER_SETTINGS.DEFAULTS.SYNC_SCOPE_MODE,
    ARCHIVE_AFTER_DAYS: settings.ARCHIVE_AFTER_DAYS || INTERIOR_MASTER_SETTINGS.DEFAULTS.ARCHIVE_AFTER_DAYS
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

function getTodoistProjectId_(useTestProject) {
  var settings = getInteriorMasterSettings_();
  var testId = (settings.TODOIST_TEST_PROJECT_ID || '').toString().trim();
  var prodId = (settings.TODOIST_PROJECT_ID || '').toString().trim();
  if (useTestProject && testId) return testId;
  return prodId;
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
