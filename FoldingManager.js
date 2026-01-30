/** FoldingManager.gs
 *
 * ✅ 버튼 1개로 “접기/펴기” 토글
 * ✅ 블록 기준(START_ROW=4, BLOCK_HEIGHT=9):
 *   - 블록 1: 4~12행, 블록 2: 13~21행 … (즉, 11행= r+7, 12행= r+8)
 * ✅ 12행(= r+8)은 어떤 상태에서도 숨기지 않음
 *
 * [접기 규칙]
 * 1) 상태가 "완료" 또는 "취소"  → 5~11행(= r+1 ~ r+7) 숨김
 * 2) 상태가 "세팅 대기" 또는 "세팅완료(에비대기)"
 *    → 5~8행(= r+1 ~ r+4) + 11행(= r+7) 숨김
 * 3) 그 외 상태(진행/실측대기/디자인 작업/엑셀 작업/대기 등) → 숨김 없음
 *
 * [토글 동작]
 * - 이미 접힌 행이 하나라도 있으면: 전체 “펼치기”
 * - 아니면: 규칙에 따라 전체 “접기”
 */

function toggleProjectRowFolding() {
  var sheet = getMainSheet_();
  var blockHeight = getBlockHeight_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.START_ROW) return;

  var stopCtl = makeStopController_();

  // 1) 지금 이미 접힌 상태인지 감지(대상 범위: 각 블록의 r+1~r+7)
  var anyFolded = false;
  for (var r = CONFIG.START_ROW; r <= lastRow; r += blockHeight) {
    if (stopCtl.check(sheet, r)) break;

    var pname = sheet.getRange(r, CONFIG.POS_NAME.col).getDisplayValue();
    if (!isValidName(pname)) continue;

    // r+1~r+7 중 하나라도 숨김이면 "접힌 상태"로 간주
    for (var rr = r + 1; rr <= r + 7; rr++) {
      if (rr > lastRow) break;
      if (sheet.isRowHiddenByUser(rr)) { anyFolded = true; break; }
    }
    if (anyFolded) break;
  }

  // 2) 토글 실행
  if (anyFolded) {
    _expandAllBlocks_(sheet, blockHeight, lastRow);
  } else {
    _foldBlocksByStatus_(sheet, blockHeight, lastRow);
  }
}

/** 전체 펼치기: 각 프로젝트 블록의 r+1~r+7을 모두 표시(12행=r+8은 건드릴 필요 없음) */
function _expandAllBlocks_(sheet, blockHeight, lastRow) {
  var stopCtl = makeStopController_();
  var intervals = [];

  for (var r = CONFIG.START_ROW; r <= lastRow; r += blockHeight) {
    if (stopCtl.check(sheet, r)) break;

    var pname = sheet.getRange(r, CONFIG.POS_NAME.col).getDisplayValue();
    if (!isValidName(pname)) continue;

    var s = r + 1;
    var e = Math.min(r + 7, lastRow);
    if (s <= e) intervals.push([s, e]);
  }

  _applyShowIntervals_(sheet, intervals);
}

/** 상태별 접기 적용 */
function _foldBlocksByStatus_(sheet, blockHeight, lastRow) {
  var stopCtl = makeStopController_();

  // 한번에 show/hide 하기 위해 interval을 모아 처리
  var toShow = [];
  var toHide = [];

  for (var r = CONFIG.START_ROW; r <= lastRow; r += blockHeight) {
    if (stopCtl.check(sheet, r)) break;

    var pname = sheet.getRange(r, CONFIG.POS_NAME.col).getDisplayValue();
    if (!isValidName(pname)) continue;

    // 기본: 우선 r+1~r+7은 모두 펼친 다음, 상태에 따라 필요한 행만 숨김
    var baseS = r + 1;
    var baseE = Math.min(r + 7, lastRow);
    if (baseS <= baseE) toShow.push([baseS, baseE]);

    var status = _getStatus_(sheet, r);

    // 1) 완료/취소 → 5~11행 숨김 (r+1~r+7)
    if (_isClosedStatus_(status)) {
      if (baseS <= baseE) toHide.push([baseS, baseE]);
      continue;
    }

    // 2) 세팅 대기 / 세팅완료(에비대기) → 5~8행 + 11행 숨김
    if (_isSettingWaitOrSettingDoneStatus_(status)) {
      // 5~8행: r+1~r+4
      var s1 = r + 1, e1 = Math.min(r + 4, lastRow);
      if (s1 <= e1) toHide.push([s1, e1]);

      // 11행: r+7 (단일)
      var row11 = r + 7;
      if (row11 <= lastRow) toHide.push([row11, row11]);
      continue;
    }

    // 그 외 상태는 숨김 없음
  }

  // 적용 순서: 먼저 show → 그 다음 hide
  _applyShowIntervals_(sheet, toShow);
  _applyHideIntervals_(sheet, toHide);
}

/** 상태 읽기: 블록 시작행의 G열(7) */
function _getStatus_(sheet, blockStartRow) {
  var v = sheet.getRange(blockStartRow, 7).getDisplayValue(); // G
  return (v || "").toString().trim();
}

/** 완료/취소 여부 */
function _isClosedStatus_(status) {
  var s = (status || "").toString().trim();
  return (s === "완료" || s === "취소");
}

/** “세팅 대기” 또는 “세팅완료(에비대기)” 여부(공백 유무 흡수) */
function _isSettingWaitOrSettingDoneStatus_(status) {
  var s = (status || "").toString().trim();
  var n = s.replace(/\s+/g, "");
  return (n === "세팅대기" || n === "세팅완료(에비대기)");
}

/** interval 유틸: show */
function _applyShowIntervals_(sheet, intervals) {
  var merged = _mergeIntervals_(intervals);
  for (var i = 0; i < merged.length; i++) {
    var s = merged[i][0], e = merged[i][1];
    if (s <= e) sheet.showRows(s, e - s + 1);
  }
}

/** interval 유틸: hide */
function _applyHideIntervals_(sheet, intervals) {
  var merged = _mergeIntervals_(intervals);
  for (var i = 0; i < merged.length; i++) {
    var s = merged[i][0], e = merged[i][1];
    if (s <= e) sheet.hideRows(s, e - s + 1);
  }
}

/** interval 병합: [[start,end], ...] */
function _mergeIntervals_(intervals) {
  if (!intervals || intervals.length === 0) return [];
  intervals.sort(function(a, b) { return a[0] - b[0]; });

  var out = [];
  var cur = [intervals[0][0], intervals[0][1]];

  for (var i = 1; i < intervals.length; i++) {
    var s = intervals[i][0], e = intervals[i][1];
    if (s <= cur[1] + 1) {
      cur[1] = Math.max(cur[1], e);
    } else {
      out.push(cur);
      cur = [s, e];
    }
  }
  out.push(cur);
  return out;
}
