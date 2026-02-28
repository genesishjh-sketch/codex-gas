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
    projects: [
      'project_code',
      'client_id',
      'client_name',
      'project_type',
      'contract_date',
      'balance_date',
      'address',
      'memo',
      'address_link',
      'folder_link',
      'before_photo_link',
      'construction_photo_link',
      'after_photo_link',
      'avi_link',
      'blog_link',
      'viewer_link',
      'edit_link',
      'sheet_link'
    ],
    milestones: ['project_code', 'section', 'step_name', 'plan_date', 'done_date', 'manager']
  }
};

var INTERIOR_SYNC_KEYS = {
  AUTO_SYNC_ON_OPEN: 'INTERIOR_SYNC_ON_OPEN'
};

var INTERIOR_SYNC_TRIGGER_HANDLER = 'runInteriorDbSyncByTrigger';

/** UI 사용 가능 여부 확인 (트리거 실행 대비) */
function getUiIfAvailable_() {
  try {
    return SpreadsheetApp.getUi();
  } catch (e) {
    return null;
  }
}

function alertIfPossible_(ui, message) {
  if (!ui || !message) return;
  ui.alert(message);
}

/**
 * (호환용) 별도 메뉴가 필요한 환경에서 사용할 수 있는 메뉴 생성 함수
 * 실제 기본 메뉴 등록은 Main.js의 onOpen()에서 처리합니다.
 */
/**
 * 메인 실행 함수
 * - Source 블록 구조를 순회하여 clients/projects UPSERT
 * - milestones는 프로젝트코드 단위로 삭제 후 재삽입
 */
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
  var normalizedCurrent = firstRowValues.map(function(v) {
    return (v || '').toString().trim();
  });

  var needsWrite = false;
  for (var i = 0; i < headers.length; i++) {
    if (normalizedCurrent[i] !== headers[i]) {
      needsWrite = true;
      break;
    }
  }

  if (!needsWrite) return;
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
 * - 링크: 링크명별 URL 분리 저장
 */
function buildRecordFromAnchor_(sourceSheet, anchorRow, nextAnchorRow) {
  var projectCode = normalizeProjectCode_(readCellDisplay_(sourceSheet, anchorRow, 2));
  var blockEndRow = nextAnchorRow ? (nextAnchorRow - 1) : Math.min(sourceSheet.getLastRow(), anchorRow + Math.max(9, getBlockHeight_(sourceSheet) + 1));
  if (!projectCode || blockEndRow < anchorRow) return null;

  var blockRows = blockEndRow - anchorRow + 1;
  var maxCols = Math.min(sourceSheet.getLastColumn(), 220);
  var displayBlock = sourceSheet.getRange(anchorRow, 1, blockRows, maxCols).getDisplayValues();
  var valueBlock = sourceSheet.getRange(anchorRow, 1, blockRows, maxCols).getValues();

  function toIdxRow(absRow) { return absRow - anchorRow; }
  function getDisplay(absRow, col) {
    var r = toIdxRow(absRow), c = col - 1;
    if (r < 0 || r >= displayBlock.length || c < 0 || c >= maxCols) return '';
    return (displayBlock[r][c] || '').toString().trim();
  }
  function getValue(absRow, col) {
    var r = toIdxRow(absRow), c = col - 1;
    if (r < 0 || r >= valueBlock.length || c < 0 || c >= maxCols) return '';
    return valueBlock[r][c];
  }
  function findRawByLabel(startRow, endRow, labelCol, valueCol, labels) {
    var wanted = {};
    (labels || []).forEach(function(label) { wanted[normalizeLinkLabel_(label)] = true; });
    for (var rr = startRow; rr <= endRow; rr++) {
      var key = normalizeLinkLabel_(getDisplay(rr, labelCol));
      if (!wanted[key]) continue;
      return getValue(rr, valueCol);
    }
    return '';
  }
  function findDisplayByLabel(startRow, endRow, labelCol, valueCol, labels) {
    var raw = findRawByLabel(startRow, endRow, labelCol, valueCol, labels);
    return (raw === null || raw === undefined) ? '' : String(raw).trim();
  }

  var baseRow = Math.min(anchorRow + 1, blockEndRow);
  var profileRow = Math.max(1, anchorRow - 6);

  // 통합관리시트 레이아웃 기준: Anchor(B열 프로젝트 코드) 기준 고객정보는 상단 고정 영역(D/C열)에 위치.
  // 예) Anchor=11행일 때 고객명=5행 D열, 연락처=6행 D열
  var clientName = readCellDisplay_(sourceSheet, profileRow, 4) || getDisplay(baseRow, 4);
  var phone = readCellDisplay_(sourceSheet, profileRow + 1, 4)
    || findDisplayByLabel(baseRow, blockEndRow, 3, 4, ['연락처', '휴대폰', '핸드폰', '전화번호', '연락처(핸드폰)', '연락처(휴대폰)']);
  var clientId = makeClientId_(clientName, phone, projectCode);

  var projectType = readCellDisplay_(sourceSheet, profileRow, 3) || getDisplay(baseRow, 3);
  var contractDate = toYmd_(findRawByLabel(baseRow, blockEndRow, 3, 4, ['계약일', '계약']));
  var balanceDate = toYmd_(findRawByLabel(baseRow, blockEndRow, 3, 4, ['잔금', '잔금일']));

  var addr1 = findDisplayByLabel(baseRow, blockEndRow, 5, 6, ['주소']);
  var addr2 = findDisplayByLabel(baseRow, blockEndRow, 5, 6, ['상세주소', '추가주소']);
  var address = [addr1, addr2].filter(function(v) { return v; }).join(' ');

  var memo = getDisplay(baseRow, 12);

  var links = extractProjectLinks_(sourceSheet, anchorRow, blockEndRow, displayBlock);

  var milestones = [];

  for (var r1 = baseRow; r1 <= blockEndRow; r1++) {
    var stepName = getDisplay(r1, 7);
    var planDate1 = toYmd_(getValue(r1, 8));
    var doneDate = toYmd_(getValue(r1, 9));

    if (stepName === '완료') {
      continue;
    }

    if (stepName || planDate1 || doneDate) {
      milestones.push([projectCode, '홈스타일링', stepName, planDate1, doneDate, '']);
    }
  }

  for (var r2 = baseRow; r2 <= blockEndRow; r2++) {
    var category = getDisplay(r2, 13);
    var planDate2 = toYmd_(getValue(r2, 14));
    var manager = getDisplay(r2, 16);

    if (planDate2) {
      milestones.push([projectCode, '시공/지원', category, planDate2, '', manager]);
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
    addressLink: links.addressLink,
    folderLink: links.folderLink,
    beforePhotoLink: links.beforePhotoLink,
    constructionPhotoLink: links.constructionPhotoLink,
    afterPhotoLink: links.afterPhotoLink,
    aviLink: links.aviLink,
    blogLink: links.blogLink,
    viewerLink: links.viewerLink,
    editLink: links.editLink,
    sheetLink: links.sheetLink,
    milestones: milestones
  };
}

/** 프로젝트 코드 후행 노이즈 제거 (예: " /", 특수기호, 공백) */
function normalizeProjectCode_(projectCode) {
  var trimmed = (projectCode || '').toString().trim();
  if (!trimmed) return '';

  // 기본 코드 패턴(YYMMDD + 본문)은 유지하고, 끝의 공백/기호만 제거.
  return trimmed
    // 일반/전각/제로폭 계열 공백 제거
    .replace(/[\s\u00A0\u2000-\u200D\u2060\uFEFF\u3000]+$/g, '')
    // 끝에 붙은 특수기호(ASCII + 전각 슬래시/역슬래시 등) 제거
    .replace(/[\/\\\|,:;~`!@#$%^&*+=<>?"'\-·•…。，、！＠＃＄％＾＆＊（）＿＋＝＜＞？／＼]+$/g, '')
    // 혹시 남은 비정상 꼬리문자(제어문자/특수기호 등) 한 번 더 제거
    .replace(/[^0-9A-Za-z가-힣()\[\]\s]+$/g, '')
    .trim();
}

function extractProjectLinks_(sourceSheet, anchorRow, blockEndRow, displayBlock) {
  var scanStartRow = Math.max(1, anchorRow);
  var scanEndRow = Math.max(scanStartRow, blockEndRow || anchorRow);
  var linkSpecs = [
    { key: 'addressLink', labels: ['주소링크', '주소 링크'], useRightCell: false, fallback: { row: anchorRow - 3, col: 6 } },
    { key: 'folderLink', labels: ['[폴더] 링크', '[폴더]링크', '[폴더]', '폴더링크', '폴더 링크'], useRightCell: false, fallback: { row: anchorRow - 7, col: 9 } },
    { key: 'beforePhotoLink', labels: ['01 Before 사진 링크', '01 Before 사진링크', '01Before사진링크', 'before 사진 링크'], useRightCell: false },
    { key: 'constructionPhotoLink', labels: ['02 시공 사진 링크', '02 시공 사진링크', '02시공사진링크', '시공 사진 링크'], useRightCell: false },
    { key: 'afterPhotoLink', labels: ['03 After 사진 링크', '03 After 사진링크', '03After사진링크', 'after 사진 링크'], useRightCell: false },
    { key: 'aviLink', labels: ['에비링크', '에비 링크'], useRightCell: false },
    { key: 'blogLink', labels: ['블로그 링크', '블로그링크'], useRightCell: false, fallback: { row: anchorRow - 7, col: 11 } },
    { key: 'viewerLink', labels: ['(뷰어) 링크', '(뷰어)링크', '뷰어 링크', '뷰어링크'], useRightCell: true },
    { key: 'editLink', labels: ['(수정) 링크', '(수정)링크', '수정 링크', '수정링크'], useRightCell: true },
    { key: 'sheetLink', labels: ['(시트) 링크', '(시트)링크', '시트 링크', '시트링크'], useRightCell: true }
  ];

  var result = {};
  linkSpecs.forEach(function(spec) {
    var found = findLinkByLabels_(sourceSheet, scanStartRow, scanEndRow, spec.labels, spec.useRightCell, anchorRow, displayBlock);
    if (!found && spec.fallback) {
      found = readCellLink_(sourceSheet, spec.fallback.row, spec.fallback.col);
    }
    result[spec.key] = found || '';
  });

  return result;
}

function findLinkByLabels_(sheet, startRow, endRow, labels, useRightCell, anchorRow, displayBlock) {
  if (!labels || labels.length === 0 || startRow > endRow) return '';

  var wanted = {};
  labels.forEach(function(label) {
    wanted[normalizeLinkLabel_(label)] = true;
  });

  var lastCol = Math.min(sheet.getLastColumn(), 220);
  for (var row = startRow; row <= endRow; row++) {
    var rowVals;
    if (displayBlock && anchorRow && row >= anchorRow && (row - anchorRow) < displayBlock.length) {
      rowVals = displayBlock[row - anchorRow];
    } else {
      rowVals = sheet.getRange(row, 1, 1, lastCol).getDisplayValues()[0];
    }

    for (var c = 0; c < Math.min(rowVals.length, lastCol); c++) {
      var label = normalizeLinkLabel_(rowVals[c]);
      if (!wanted[label]) continue;

      var baseCol = c + 1;
      var colCandidates = useRightCell
        ? [baseCol + 1, baseCol, baseCol - 1]
        : [baseCol, baseCol + 1];

      for (var i = 0; i < colCandidates.length; i++) {
        var url = readCellLink_(sheet, row, colCandidates[i]);
        if (url) return url;
      }
    }
  }

  return '';
}

function normalizeLinkLabel_(value) {
  return (value || '').toString().replace(/\s+/g, '').toLowerCase();
}

function readCellLink_(sheet, row, col) {
  if (row < 1 || col < 1) return '';
  var cell = sheet.getRange(row, col);
  var url = getUrlFromCell_(cell);
  if (url) return url;
  var display = (cell.getDisplayValue() || '').toString().trim();
  return (display.indexOf('http') === 0) ? display : '';
}

/** clients/projects 공통 UPSERT (헤더 제외, 2행부터 반영) */
/**
 * milestones 갱신
 * - 대상 프로젝트코드들의 기존 행을 삭제
 * - 새 milestones 행 삽입
 */
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
  var normalized = normalizeProjectCode_(projectCode);
  if (!normalized) return false;

  // 날짜 6자리로 시작하면 유효로 판단하고, 뒤 텍스트는 유연하게 허용
  // (예: 괄호 뒤 특수문자/공백이 있어도 정규화 후 통과)
  var pattern = /^\d{6}(?:\s+.+)?$/;
  return pattern.test(normalized);
}

/** 프로젝트 코드 후보 검사: 날짜 6자리로 시작하는지 */
function isProjectCodeCandidate_(projectCode) {
  var trimmed = (projectCode || '').toString().trim();
  if (!trimmed) return false;
  return /^\d{6}/.test(trimmed);
}

/** 고객 ID 형식 검사: 고객명(필수) + 연락처4자리(선택) */
function isValidClientIdFormat_(clientId) {
  var trimmed = (clientId || '').toString().trim();
  if (!trimmed) return false;
  var pattern = /^[^\d\s]+(?:\d{4})?$/;
  return pattern.test(trimmed);
}

/** 고객ID 생성: 고객명 + 연락처 마지막 4자리 숫자(없으면 고객명만 사용) */
function makeClientId_(name, phone, projectCode) {
  var safeName = normalizeClientName_(name);
  var digits = (phone || '').toString().replace(/\D/g, '');
  var last4 = digits ? digits.slice(-4) : '';
  if (safeName && last4) return safeName + last4;
  if (safeName && digits) return safeName + digits.slice(-Math.min(4, digits.length));
  if (safeName) return safeName;

  // 이름까지 비어 있으면 최소한 프로젝트코드 기반으로 키를 만든다.
  var normalizedCode = normalizeProjectCode_(projectCode);
  if (normalizedCode) return normalizedCode.replace(/\s+/g, '');

  return '';
}

function normalizeClientName_(name) {
  return (name || '').toString().replace(/\s+/g, '').trim();
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

function findValueByLabel_(sheet, startRow, endRow, labelCol, valueCol, labels) {
  var raw = findValueByLabelRaw_(sheet, startRow, endRow, labelCol, valueCol, labels);
  return (raw === null || raw === undefined) ? '' : String(raw).trim();
}

function findValueByLabelRaw_(sheet, startRow, endRow, labelCol, valueCol, labels) {
  if (startRow < 1 || endRow < startRow) return '';

  var wanted = {};
  (labels || []).forEach(function(label) {
    wanted[normalizeLinkLabel_(label)] = true;
  });

  var rows = endRow - startRow + 1;
  var labelVals = sheet.getRange(startRow, labelCol, rows, 1).getDisplayValues();
  for (var i = 0; i < labelVals.length; i++) {
    var key = normalizeLinkLabel_(labelVals[i][0]);
    if (!wanted[key]) continue;
    return readCellValue_(sheet, startRow + i, valueCol);
  }

  return '';
}

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
