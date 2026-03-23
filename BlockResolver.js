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

  return null;
}
