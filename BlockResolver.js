/** BlockResolver.gs
 * 방금 생성된 블록 1개를 찾는 로직.
 */

function resolveTargetBlockStartRowForClickUp_(summary, settings) {
  var basic = (settings && settings.basic) ? settings.basic : {};
  var blockStart = Number(basic.BLOCK_START_ROW || 4);
  var blockHeight = Number(basic.BLOCK_HEIGHT || 9);

  // 1) 코어가 완료 블록을 알려준 경우: 마지막 1개 사용
  var completedRows = (summary && summary.completedRows) ? summary.completedRows.slice() : [];
  if (completedRows.length > 0) return completedRows[completedRows.length - 1];

  // 2) 현재 활성 셀 기준 블록 추정
  var activeRange = SpreadsheetApp.getActiveRange();
  if (activeRange) {
    var row = activeRange.getRow();
    if (row >= blockStart) {
      var normalized = blockStart + Math.floor((row - blockStart) / blockHeight) * blockHeight;
      return normalized;
    }
  }

  // 3) pendingRows가 있으면 마지막 행
  var pendingRows = (summary && summary.pendingRows) ? summary.pendingRows.slice() : [];
  if (pendingRows.length > 0) return pendingRows[pendingRows.length - 1];

  // 4) 마지막 fallback: 원본 시트에서 "아직 task_id 없는 최신 블록" 1개 탐색
  var latest = findLatestUnsyncedBlockStartRow_(settings);
  if (latest) return latest;

  return null;
}

function findLatestUnsyncedBlockStartRow_(settings) {
  try {
    var basic = (settings && settings.basic) ? settings.basic : {};
    var sheetName = stringValue_(basic.SOURCE_SHEET_NAME || (typeof CONFIG !== 'undefined' ? CONFIG.SHEET_NAME : '통합관리시트'));
    var blockStart = Number(basic.BLOCK_START_ROW || 4);
    var blockHeight = Number(basic.BLOCK_HEIGHT || 9);
    var repOffset = Number(basic.REPRESENTATIVE_ROW_OFFSET || 1);
    var taskIdCol = colToNumber_(basic.CLICKUP_TASK_ID_COLUMN || 'U');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(sheetName);
    if (!sh) return null;

    var lastRow = sh.getLastRow();
    if (lastRow < blockStart) return null;

    var top = blockStart + Math.floor((lastRow - blockStart) / blockHeight) * blockHeight;
    for (var r = top; r >= blockStart; r -= blockHeight) {
      var nameVal = sh.getRange(r + (CONFIG.POS_NAME.row || 0), CONFIG.POS_NAME.col || 3).getDisplayValue();
      if (!isValidName(nameVal)) continue;
      if (isClosedBlock_(sh, r)) continue;

      var repRow = r + repOffset;
      var taskId = sh.getRange(repRow, taskIdCol).getDisplayValue();
      if (stringValue_(taskId)) continue;

      return r;
    }
  } catch (e) {}

  return null;
}
