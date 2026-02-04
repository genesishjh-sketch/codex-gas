/** Utils.gs
 * 공통 유틸리티 모음
 * - 블록(프로젝트) 탐색/판정
 * - URL/ID 추출
 * - 시트/로그 시트 생성
 * - 주소/전화번호 정규화
 */

/** =========================
 *  프로젝트명(블록 시작) 판정
 *  =========================
 *  - "멱살반/반멱살/스타일링대행"으로 시작
 *  - "님" 포함
 *  - 괄호(동네) 표기는 있으면 좋지만, 없다고 무조건 제외하지 않음
 */
function isValidName(nameVal) {
  if (nameVal === null || nameVal === undefined) return false;
  var s = String(nameVal).trim();
  if (!s || s === "#N/A") return false;

  var validation = (typeof CONFIG !== "undefined" && CONFIG.NAME_VALIDATION) ? CONFIG.NAME_VALIDATION : {};
  if (validation && validation.allowAny) return true;

  var suffix = (validation && validation.suffix) ? validation.suffix : "님";
  var requireSuffix = !!(validation && validation.requireSuffix);
  var prefixes = (validation && validation.prefixes) ? validation.prefixes : [];

  var hasPrefix = false;
  if (prefixes && prefixes.length > 0) {
    for (var i = 0; i < prefixes.length; i++) {
      var p = String(prefixes[i] || "").trim();
      if (!p) continue;
      if (s.indexOf(p) === 0) { hasPrefix = true; break; }
    }
  } else {
    hasPrefix = /^(멱살반|반멱살|스타일링대행|단기)\b/.test(s);
  }

  var hasSuffix = suffix ? (s.indexOf(suffix) !== -1) : false;

  if (requireSuffix && !hasSuffix) return false;
  if (!hasPrefix && !hasSuffix) return false;

  return true;
}

/** =========================
 *  URL에서 ID 추출
 *  ========================= */
function extractIdFromUrl(url) {
  if (!url) return "";
  var s = String(url).trim();

  // Drive folder
  var m = s.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (m && m[1]) return m[1];

  // Drive/Docs file (spreadsheets/docs/presentation etc.)
  m = s.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m && m[1]) return m[1];

  // open?id=...
  m = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m && m[1]) return m[1];

  // fallback: 알 수 없으면 원문 반환(빈값 방지)
  return s;
}

/** =========================
 *  시트 접근/생성
 *  ========================= */
function getMainSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) throw new Error("시트를 찾을 수 없습니다: " + CONFIG.SHEET_NAME);
  return sh;
}

function getStartRow_() {
  return (CONFIG && CONFIG.START_ROW) ? CONFIG.START_ROW : 4;
}

/** =========================
 *  잔금일 추출
 *  - D9(블록 기준) 셀에서 날짜 추출
 *  ========================= */
function extractBalanceDate_(sheet, blockStartRow) {
  var cfg = (CONFIG && CONFIG.BALANCE_DATE_CELL) ? CONFIG.BALANCE_DATE_CELL : null;
  var rowOffset = cfg && typeof cfg.row === "number" ? cfg.row : 5;
  var col = cfg && typeof cfg.col === "number" ? cfg.col : 4;
  var cell = sheet.getRange(blockStartRow + rowOffset, col).getValue();
  if (cell instanceof Date) return new Date(cell.getTime());
  if (typeof cell === "string" && cell.trim() !== "") {
    var parsed = new Date(cell);
    if (!isNaN(parsed.getTime())) return parsed;
  }
  return null;
}

/** =========================
 *  잔금 상태 셀(D5) 값 확인
 *  ========================= */
function getBalanceStatusValue_(sheet, blockStartRow) {
  var cfg = (CONFIG && CONFIG.BALANCE_STATUS_CELL) ? CONFIG.BALANCE_STATUS_CELL : null;
  var rowOffset = cfg && typeof cfg.row === "number" ? cfg.row : 1;
  var col = cfg && typeof cfg.col === "number" ? cfg.col : 4;
  return sheet.getRange(blockStartRow + rowOffset, col).getDisplayValue();
}

/**
 * 시트 없으면 만들고, 헤더 없으면 1행에 헤더 세팅
 * @param {string} sheetName
 * @param {string[]} headers
 */
function ensureSheet_(sheetName, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);

  if (headers && headers.length > 0) {
    var width = headers.length;
    var first = sh.getRange(1, 1, 1, width).getValues()[0];
    var empty = first.every(function(v){ return v === "" || v === null; });

    if (sh.getLastRow() === 0 || empty) {
      sh.getRange(1, 1, 1, width).setValues([headers]).setFontWeight("bold");
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

/** =========================
 *  블록 높이(9행) 자동 탐지
 *  ========================= */
function detectBlockHeight_(sheet) {
  var startRow = getStartRow_();
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow + 10) return CONFIG.BLOCK_HEIGHT;

  var checkRows = Math.min(lastRow - startRow + 1, 500);
  var names = sheet.getRange(startRow, CONFIG.POS_NAME.col, checkRows, 1).getDisplayValues();

  var starts = [];
  for (var i = 0; i < names.length; i++) {
    if (isValidName(names[i][0])) starts.push(i);
  }
  if (starts.length < 2) return CONFIG.BLOCK_HEIGHT;

  var freq = {};
  for (var k = 1; k < starts.length; k++) {
    var d = starts[k] - starts[k - 1];
    if (d >= 4 && d <= 30) freq[d] = (freq[d] || 0) + 1;
  }

  var best = CONFIG.BLOCK_HEIGHT, bestCnt = -1;
  Object.keys(freq).forEach(function(key) {
    if (freq[key] > bestCnt) { bestCnt = freq[key]; best = Number(key); }
  });

  return best || CONFIG.BLOCK_HEIGHT;
}

function getBlockHeight_(sheet) {
  return (CONFIG && CONFIG.BLOCK_HEIGHT) ? CONFIG.BLOCK_HEIGHT : detectBlockHeight_(sheet);
}

function getBlockStartRow_(row, blockHeight) {
  var start = getStartRow_();
  if (row < start) return null;
  return start + Math.floor((row - start) / blockHeight) * blockHeight;
}

/** =========================
 *  전화번호 정규화
 *  ========================= */
function normalizePhone_(raw) {
  if (!raw) return "";
  var s = String(raw).replace(/[^\d]/g, "");

  // 010-xxxx-xxxx
  if (s.length === 11 && s.indexOf("010") === 0) {
    return s.slice(0,3) + "-" + s.slice(3,7) + "-" + s.slice(7);
  }
  // 010-xxx-xxxx
  if (s.length === 10 && s.indexOf("010") === 0) {
    return s.slice(0,3) + "-" + s.slice(3,6) + "-" + s.slice(6);
  }
  return String(raw).trim();
}

/** =========================
 *  완료/취소 블록 판정
 *  - 블록 시작행 G열 값이 "완료" 또는 "취소"면 true
 *  ========================= */
function isClosedBlock_(sheet, blockStartRow) {
  try {
    var v = sheet.getRange(blockStartRow, 7).getDisplayValue(); // G
    v = String(v || "").trim();
    if (v === "완료" || v === "취소") return true;
    if (typeof findStatusInBlock_ === "function") {
      var status = findStatusInBlock_(sheet, blockStartRow);
      return (status === "완료" || status === "취소");
    }
    return false;
  } catch (e) {
    return false;
  }
}

/** =========================
 *  종료 조건 컨트롤러
 *  - (B:No, C:프로젝트명) 모두 빈 블록이 연속 N개면 중단
 *  ========================= */
function makeStopController_() {
  var streak = 0;
  var startRow = getStartRow_();

  return {
    check: function(sheet, blockStartRow) {
      // blockStartRow는 보통 START_ROW부터 blockHeight씩 증가
      if (blockStartRow < startRow) return false;

      var noVal = sheet.getRange(blockStartRow, CONFIG.POS_NO.col).getDisplayValue();
      var nameVal = sheet.getRange(blockStartRow, CONFIG.POS_NAME.col).getDisplayValue();

      var noEmpty = String(noVal || "").trim() === "";
      var nameEmpty = String(nameVal || "").trim() === "";

      if (noEmpty && nameEmpty) streak++;
      else streak = 0;

      return streak >= (CONFIG.STOP_AFTER_EMPTY_BLOCKS || 3);
    }
  };
}

/** =========================
 *  주소 문자열 분리
 *  - base: 지번까지 ("... 719-8")
 *  - extra: 지번 뒤 나머지(호/층 등) — 괄호(도로명) 제거
 *  ========================= */
function splitAddressExtra_(raw) {
  var s = String(raw || "");
  s = s.replace(/\r\n/g, "\n").replace(/\n+/g, " ").trim();
  s = s.replace(/\s+/g, " ").trim();
  if (!s) return { base: "", extra: "" };

  // 첫 지번(숫자-숫자)까지만 base
  var m = s.match(/^(.+?\d+(?:-\d+)?)(.*)$/);
  if (!m) return { base: s, extra: "" };

  var base = String(m[1] || "").trim();
  var rest = String(m[2] || "").trim();

  // 괄호 덩어리 제거 → extra만 남김
  rest = rest.replace(/\([^)]*\)/g, " ").replace(/\s+/g, " ").trim();

  return { base: base, extra: rest };
}

/** 카카오 결과 정리용: 주소 앞쪽 광역단위 제거 */
function cleanPrefix_(text) {
  if (!text) return "";
  return String(text).replace(
    /^(서울(특별시)?|경기(도)?|인천(광역시)?|대구(광역시)?|부산(광역시)?|광주(광역시)?|대전(광역시)?|울산(광역시)?|제주(특별자치도)?|강원(도)?|충청[남북]도|전라[남북]도|경상[남북]도)\s+/,
    ""
  ).trim();
}

/** (선택) 리치텍스트/표시값에서 URL 최대 복원 */
function getUrlFromCell_(cell) {
  try {
    var rich = cell.getRichTextValue();
    if (rich && rich.getLinkUrl()) return rich.getLinkUrl();
  } catch (e) {}

  try {
    var f = cell.getFormula();
    if (f && f.toUpperCase().indexOf("HYPERLINK(") === 0) {
      var m = f.match(/HYPERLINK\(\s*\"([^\"]+)\"/i);
      if (m && m[1]) return m[1];
    }
  } catch (e2) {}

  var v = String(cell.getDisplayValue() || "").trim();
  if (v.indexOf("http") === 0) return v;
  return "";
}
