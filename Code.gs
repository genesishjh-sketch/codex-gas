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
  TARGET_MILESTONES: 'milestones'
};

/** 스프레드시트 열기 시 사용자 메뉴 생성 */
function addInteriorSyncMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('인테리어 관리')
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
    var sourceSheet = ss.getSheetByName(INTERIOR_SYNC_CONFIG.SOURCE_SHEET);
    var clientsSheet = ss.getSheetByName(INTERIOR_SYNC_CONFIG.TARGET_CLIENTS);
    var projectsSheet = ss.getSheetByName(INTERIOR_SYNC_CONFIG.TARGET_PROJECTS);
    var milestonesSheet = ss.getSheetByName(INTERIOR_SYNC_CONFIG.TARGET_MILESTONES);

    if (!sourceSheet || !clientsSheet || !projectsSheet || !milestonesSheet) {
      throw new Error('필수 시트(통합관리시트/clients/projects/milestones) 중 일부를 찾을 수 없습니다.');
    }

    var anchors = collectAnchorRows_(sourceSheet);
    if (anchors.length === 0) {
      ss.toast('동기화할 프로젝트 코드가 없습니다.', '인테리어 관리', 5);
      return;
    }

    var clientsRows = [];
    var projectsRows = [];
    var milestonesRows = [];
    var projectCodesToRefresh = {};

    anchors.forEach(function(anchorRow) {
      var record = buildRecordFromAnchor_(sourceSheet, anchorRow);
      if (!record.projectCode) return;

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

    upsertByKey_(clientsSheet, clientsRows, 1);
    upsertByKey_(projectsSheet, projectsRows, 1);

    var targetProjectCodes = Object.keys(projectCodesToRefresh);
    replaceMilestonesByProjectCodes_(milestonesSheet, targetProjectCodes, milestonesRows);

    var doneMessage = '동기화가 완료되었습니다.\n'
      + '- clients: ' + clientsRows.length + '건 반영\n'
      + '- projects: ' + projectsRows.length + '건 반영\n'
      + '- milestones: ' + milestonesRows.length + '건 반영';

    ss.toast('동기화가 완료되었습니다.', '인테리어 관리', 5);
    ui.alert(doneMessage);
  } catch (err) {
    ui.alert('동기화 중 오류가 발생했습니다.\n' + err.message);
    throw err;
  }
}

/** B열을 순회하여 Anchor(프로젝트 코드 존재 행) 수집 */
function collectAnchorRows_(sourceSheet) {
  var lastRow = sourceSheet.getLastRow();
  if (lastRow < 1) return [];

  var colBValues = sourceSheet.getRange(1, 2, lastRow, 1).getDisplayValues();
  var anchors = [];

  for (var r = 1; r <= colBValues.length; r++) {
    var projectCode = (colBValues[r - 1][0] || '').toString().trim();
    if (projectCode) anchors.push(r);
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
  var contractDate = toYmd_(sourceSheet.getRange(anchorRow - 3, 4).getValue());
  var balanceDate = toYmd_(sourceSheet.getRange(anchorRow - 2, 4).getValue());

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
    var planDate1 = toYmd_(sourceSheet.getRange(r1, 8).getValue());
    var doneDate = toYmd_(sourceSheet.getRange(r1, 9).getValue());

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
    var planDate2 = toYmd_(sourceSheet.getRange(r2, 14).getValue());
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
