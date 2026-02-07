/**
 * 인테리어 통합관리시트 → DB 시트 동기화 스크립트
 *
 * [중요]
 * - 기준(Anchor) 행은 B열에 프로젝트 코드가 있는 행입니다.
 * - 각 데이터는 Anchor 행 기준 상대 오프셋으로 읽습니다.
 *   예) Anchor가 11행이면 고객명은 (11 - 6)행의 D열
 */

var INTERIOR_SYNC_CONFIG = {
  SOURCE_SHEET: '통합관리시트',
  TARGET_CLIENTS: 'clients',
  TARGET_PROJECTS: 'projects',
  TARGET_MILESTONES: 'milestones',
  SOURCE_SHEET_ALIASES: ['통합관리시트', '통합 관리시트'],
  TARGET_CLIENTS_ALIASES: ['clients', 'Clients', '고객', '고객DB'],
  TARGET_PROJECTS_ALIASES: ['projects', 'Projects', '프로젝트', '프로젝트DB'],
  TARGET_MILESTONES_ALIASES: ['milestones', 'Milestones', '마일스톤', '일정'],
  TARGET_HEADERS: {
    clients: ['client_id', 'client_name', 'phone'],
    projects: ['project_code', 'client_id', 'project_type', 'contract_date', 'balance_date', 'address', 'memo', 'links'],
    milestones: ['project_code', 'section', 'step_name', 'plan_date', 'done_date', 'manager']
  }
};

/**
 * (호환용) 별도 메뉴가 필요한 환경에서 사용할 수 있는 메뉴 생성 함수
 * 실제 기본 메뉴 등록은 Main.js의 onOpen()에서 처리합니다.
 */
function addInteriorSyncMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('🛋️ 인테리어 관리')
    .addItem('DB 동기화 실행', 'runInteriorDbSync')
    .addToUi();
}

/**
 * 메인 실행 함수
 * - Source 블록 구조를 순회하여 clients/projects UPSERT
 * - milestones는 프로젝트코드 단위로 삭제 후 재삽입
 */
function runInteriorDbSync() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  try {
    var sourceSheet = getSheetByAliases_(ss, INTERIOR_SYNC_CONFIG.SOURCE_SHEET_ALIASES);
    var clientsSheet = getOrCreateTargetSheet_(ss, INTERIOR_SYNC_CONFIG.TARGET_CLIENTS_ALIASES, INTERIOR_SYNC_CONFIG.TARGET_CLIENTS, INTERIOR_SYNC_CONFIG.TARGET_HEADERS.clients);
    var projectsSheet = getOrCreateTargetSheet_(ss, INTERIOR_SYNC_CONFIG.TARGET_PROJECTS_ALIASES, INTERIOR_SYNC_CONFIG.TARGET_PROJECTS, INTERIOR_SYNC_CONFIG.TARGET_HEADERS.projects);
    var milestonesSheet = getOrCreateTargetSheet_(ss, INTERIOR_SYNC_CONFIG.TARGET_MILESTONES_ALIASES, INTERIOR_SYNC_CONFIG.TARGET_MILESTONES, INTERIOR_SYNC_CONFIG.TARGET_HEADERS.milestones);

    if (!sourceSheet || !clientsSheet || !projectsSheet || !milestonesSheet) {
      var missing = [];
      if (!sourceSheet) missing.push('통합관리시트');
      if (!clientsSheet) missing.push('clients');
      if (!projectsSheet) missing.push('projects');
      if (!milestonesSheet) missing.push('milestones');

      var existingSheetNames = ss.getSheets().map(function(sheet) {
        return sheet.getName();
      }).join(', ');

      throw new Error(
        '필수 시트를 찾을 수 없습니다. 누락: '
        + missing.join('/')
        + '\n현재 스프레드시트 탭: '
        + existingSheetNames
      );
    }

    var anchors = collectAnchorRows_(sourceSheet);
    if (anchors.length === 0) {
      ss.toast('동기화할 프로젝트 코드가 없습니다.', '🛋️ 인테리어 관리', 5);
      return;
    }

    var clientsRows = [];
    var projectsRows = [];
    var milestonesRows = [];
    var projectCodesToRefresh = {};
    var invalidRecords = [];

    anchors.forEach(function(anchorRow) {
      var record = buildRecordFromAnchor_(sourceSheet, anchorRow);
      if (!record.projectCode) return;

      if (!isValidProjectCodeFormat_(record.projectCode) || !isValidClientIdFormat_(record.clientId)) {
        invalidRecords.push({
          row: anchorRow,
          projectCode: record.projectCode,
          clientId: record.clientId
        });
        return;
      }

      clientsRows.push([record.clientId, record.clientName, record.phone]);
      projectsRows.push([
        record.projectCode,
        record.clientId,
        record.projectType,
        record.contractDate,
        record.balanceDate,
        record.address,
        record.memo,
        record.links
      ]);

      projectCodesToRefresh[record.projectCode] = true;
      Array.prototype.push.apply(milestonesRows, record.milestones);
    });

    if (invalidRecords.length > 0) {
      var invalidDetails = invalidRecords.map(function(record) {
        return '행 ' + record.row + ': ' + record.projectCode + ' / ' + record.clientId;
      }).join('\n');

      ui.alert(
        '프로젝트 코드 또는 고객 ID 형식이 올바르지 않아 동기화를 중단했습니다.\n'
        + '예시) 250831 멱살반 양수정님 (성산동) / 양수정7864\n'
        + invalidDetails
      );
      return;
    }

    upsertByKey_(clientsSheet, clientsRows, 1);
    upsertByKey_(projectsSheet, projectsRows, 1);

    var targetProjectCodes = Object.keys(projectCodesToRefresh);
    replaceMilestonesByProjectCodes_(milestonesSheet, targetProjectCodes, milestonesRows);

    var doneMessage = '동기화가 완료되었습니다.\n'
      + '- clients: ' + clientsRows.length + '건 반영\n'
      + '- projects: ' + projectsRows.length + '건 반영\n'
      + '- milestones: ' + milestonesRows.length + '건 반영';

    ss.toast('동기화가 완료되었습니다.', '🛋️ 인테리어 관리', 5);
    ui.alert(doneMessage);
  } catch (err) {
    ui.alert('동기화 중 오류가 발생했습니다.\n' + err.message);
    throw err;
  }
}

/** 대상 시트가 없으면 자동 생성하고 헤더를 준비합니다. */
function getOrCreateTargetSheet_(ss, aliases, defaultName, headers) {
  var sheet = getSheetByAliases_(ss, aliases);
  if (!sheet) {
    sheet = ss.insertSheet(defaultName);
  }

  ensureHeaderRow_(sheet, headers || []);
  return sheet;
}

/** 헤더가 비어 있으면 1행에 헤더를 입력합니다. */
function ensureHeaderRow_(sheet, headers) {
  if (!sheet || !headers || headers.length === 0) return;

  var maxCols = Math.max(sheet.getMaxColumns(), headers.length);
  var headerRange = sheet.getRange(1, 1, 1, maxCols);
  var firstRowValues = headerRange.getDisplayValues()[0];
  var hasAnyValue = firstRowValues.some(function(v) {
    return (v || '').toString().trim() !== '';
  });

  if (hasAnyValue) return;

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

/** B열을 순회하여 Anchor(프로젝트 코드 존재 행) 수집 */
function collectAnchorRows_(sourceSheet) {
  var lastRow = sourceSheet.getLastRow();
  if (lastRow < 1) return [];

  var colBValues = sourceSheet.getRange(1, 2, lastRow, 1).getDisplayValues();
  var anchors = [];

  for (var r = 1; r <= colBValues.length; r++) {
    var projectCode = (colBValues[r - 1][0] || '').toString().trim();
    if (projectCode && isProjectCodeCandidate_(projectCode)) anchors.push(r);
  }
  return anchors;
}

/**
 * Anchor 행 기준 상대 오프셋으로 단일 프로젝트 레코드 구성
 *
 * 상대 위치 규칙(Anchor = a)
 * - 고객명: D(a-6)
 * - 연락처: D(a-5)
 * - 프로젝트유형: C(a-6)
 * - 계약일: D(a-3)
 * - 잔금일: D(a-2)
 * - 주소: F(a-6) + ' ' + F(a-5)
 * - 메모: E(a-1)
 * - 링크: F(a-3), I(a-7), K(a-7) 등 결합
 */
function buildRecordFromAnchor_(sourceSheet, anchorRow) {
  var projectCode = readCellDisplay_(sourceSheet, anchorRow, 2);
  var clientName = readCellDisplay_(sourceSheet, anchorRow - 6, 4);
  var phone = readCellDisplay_(sourceSheet, anchorRow - 5, 4);
  var clientId = makeClientId_(clientName, phone);

  var projectType = readCellDisplay_(sourceSheet, anchorRow - 6, 3);
  var contractDate = toYmd_(readCellValue_(sourceSheet, anchorRow - 3, 4));
  var balanceDate = toYmd_(readCellValue_(sourceSheet, anchorRow - 2, 4));

  var addr1 = readCellDisplay_(sourceSheet, anchorRow - 6, 6);
  var addr2 = readCellDisplay_(sourceSheet, anchorRow - 5, 6);
  var address = [addr1, addr2].filter(function(v) { return v; }).join(' ');

  var memo = readCellDisplay_(sourceSheet, anchorRow - 1, 5);

  var links = [
    readCellDisplay_(sourceSheet, anchorRow - 3, 6),
    readCellDisplay_(sourceSheet, anchorRow - 7, 9),
    readCellDisplay_(sourceSheet, anchorRow - 7, 11)
  ].filter(function(v) { return v; }).join('\n');

  var milestones = [];

  // 섹션1) 홈스타일링 일정: G~I, (a-6) ~ (a-2)
  for (var r1 = anchorRow - 6; r1 <= anchorRow - 2; r1++) {
    if (r1 < 1) continue;

    var stepName = readCellDisplay_(sourceSheet, r1, 7);
    var planDate1 = toYmd_(readCellValue_(sourceSheet, r1, 8));
    var doneDate = toYmd_(readCellValue_(sourceSheet, r1, 9));

    if (stepName || planDate1 || doneDate) {
      milestones.push([
        projectCode,
        '홈스타일링',
        stepName,
        planDate1,
        doneDate,
        ''
      ]);
    }
  }

  // 섹션2) 시공/지원 일정: M~P, (a-6) ~ (a-1), N열(계획일) 필수
  for (var r2 = anchorRow - 6; r2 <= anchorRow - 1; r2++) {
    if (r2 < 1) continue;

    var category = readCellDisplay_(sourceSheet, r2, 13);
    var planDate2 = toYmd_(readCellValue_(sourceSheet, r2, 14));
    var manager = readCellDisplay_(sourceSheet, r2, 16);

    if (planDate2) {
      milestones.push([
        projectCode,
        '시공/지원',
        category,
        planDate2,
        '',
        manager
      ]);
    }
  }

  return {
    projectCode: projectCode,
    clientId: clientId,
    clientName: clientName,
    phone: phone,
    projectType: projectType,
    contractDate: contractDate,
    balanceDate: balanceDate,
    address: address,
    memo: memo,
    links: links,
    milestones: milestones
  };
}

/** clients/projects 공통 UPSERT (헤더 제외, 2행부터 반영) */
function upsertByKey_(targetSheet, rows, keyColIndex1Based) {
  if (!rows || rows.length === 0) return;

  var dataStartRow = 2;
  var lastRow = targetSheet.getLastRow();
  var keyToRowMap = {};

  if (lastRow >= dataStartRow) {
    var existingValues = targetSheet.getRange(dataStartRow, 1, lastRow - 1, targetSheet.getLastColumn()).getValues();
    for (var i = 0; i < existingValues.length; i++) {
      var key = (existingValues[i][keyColIndex1Based - 1] || '').toString().trim();
      if (key) keyToRowMap[key] = dataStartRow + i;
    }
  }

  var appendRows = [];

  rows.forEach(function(row) {
    var key = (row[keyColIndex1Based - 1] || '').toString().trim();
    if (!key) return;

    if (keyToRowMap[key]) {
      targetSheet.getRange(keyToRowMap[key], 1, 1, row.length).setValues([row]);
    } else {
      appendRows.push(row);
    }
  });

  if (appendRows.length > 0) {
    var appendStart = targetSheet.getLastRow() + 1;
    targetSheet.getRange(appendStart, 1, appendRows.length, appendRows[0].length).setValues(appendRows);
  }
}

/**
 * milestones 갱신
 * - 대상 프로젝트코드들의 기존 행을 삭제
 * - 새 milestones 행 삽입
 */
function replaceMilestonesByProjectCodes_(milestonesSheet, projectCodes, newRows) {
  var dataStartRow = 2;
  var lastRow = milestonesSheet.getLastRow();
  var codeMap = {};

  projectCodes.forEach(function(code) {
    if (code) codeMap[code] = true;
  });

  if (lastRow >= dataStartRow) {
    var rangeRows = lastRow - 1;
    var existing = milestonesSheet.getRange(dataStartRow, 1, rangeRows, 1).getDisplayValues();

    // 행 삭제는 아래에서 위로 해야 인덱스 변동 문제를 피할 수 있습니다.
    for (var i = existing.length - 1; i >= 0; i--) {
      var code = (existing[i][0] || '').toString().trim();
      if (codeMap[code]) {
        milestonesSheet.deleteRow(dataStartRow + i);
      }
    }
  }

  if (newRows && newRows.length > 0) {
    var appendStart = milestonesSheet.getLastRow() + 1;
    milestonesSheet.getRange(appendStart, 1, newRows.length, newRows[0].length).setValues(newRows);
  }
}

/**
 * 통합관리시트에 체크박스 실행 버튼을 생성합니다.
 * - A1: 체크박스(실행 스위치)
 * - B1: 안내 문구
 */
/** 별칭 목록 기준으로 시트를 조회합니다. (정확 일치 우선, 대소문자 무시 보조) */
function getSheetByAliases_(ss, aliases) {
  if (!ss || !aliases || aliases.length === 0) return null;

  for (var i = 0; i < aliases.length; i++) {
    var exact = ss.getSheetByName(aliases[i]);
    if (exact) return exact;
  }

  var normalizedAliasMap = {};
  for (var j = 0; j < aliases.length; j++) {
    normalizedAliasMap[(aliases[j] || '').toString().trim().toLowerCase()] = true;
  }

  var sheets = ss.getSheets();
  for (var k = 0; k < sheets.length; k++) {
    var normalizedSheetName = (sheets[k].getName() || '').toString().trim().toLowerCase();
    if (normalizedAliasMap[normalizedSheetName]) return sheets[k];
  }

  return null;
}

/** 프로젝트 코드 형식 검사: "YYMMDD ... ...님 (지역)" */
function isValidProjectCodeFormat_(projectCode) {
  var trimmed = (projectCode || '').toString().trim();
  if (!trimmed) return false;
  var pattern = /^\d{6}\s+.+\s+.+님\s+\(.+\)$/;
  return pattern.test(trimmed);
}

/** 프로젝트 코드 후보 검사: 날짜 6자리로 시작하는지 */
function isProjectCodeCandidate_(projectCode) {
  var trimmed = (projectCode || '').toString().trim();
  if (!trimmed) return false;
  return /^\d{6}/.test(trimmed);
}

/** 고객 ID 형식 검사: "이름+4자리숫자" */
function isValidClientIdFormat_(clientId) {
  var trimmed = (clientId || '').toString().trim();
  if (!trimmed) return false;
  var pattern = /^[^\d\s]+\d{4}$/;
  return pattern.test(trimmed);
}

/** 고객ID 생성: 고객명 + 연락처 마지막 4자리 숫자 */
function makeClientId_(name, phone) {
  var safeName = (name || '').toString().trim();
  var digits = (phone || '').toString().replace(/\D/g, '');
  var last4 = digits ? digits.slice(-4) : '';
  return safeName + last4;
}

/** 셀 표시값 읽기 (행/열 유효성 보호) */
function readCellDisplay_(sheet, row, col) {
  if (row < 1 || col < 1) return '';
  return (sheet.getRange(row, col).getDisplayValue() || '').toString().trim();
}

/** 셀 원본값 읽기 (행/열 유효성 보호) */
function readCellValue_(sheet, row, col) {
  if (row < 1 || col < 1) return '';
  return sheet.getRange(row, col).getValue();
}

/** 날짜/문자열을 YYYY-MM-DD 문자열로 통일 */
function toYmd_(value) {
  if (!value) return '';
  var tz = Session.getScriptTimeZone() || 'Asia/Seoul';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  }

  // 이미 텍스트인 경우에도 Date 변환이 가능하면 동일 포맷으로 반환
  var maybeDate = new Date(value);
  if (!isNaN(maybeDate.getTime())) {
    return Utilities.formatDate(maybeDate, tz, 'yyyy-MM-dd');
  }

  return value.toString().trim();
}
