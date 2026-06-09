/** CompletedProjectMover.gs
 * 통합관리시트에서 G열 상태가 "완료"인 프로젝트 블록을 완료 탭으로 이동한다.
 */

function runMoveCompletedProjects() {
  var ui = SpreadsheetApp.getUi();
  var completedSheetName = getCompletedSheetName_();

  var response = ui.alert(
    "✅ 완료 프로젝트 이동",
    "통합관리시트에서 G열 상태가 '완료'인 프로젝트를 '" + completedSheetName + "' 탭으로 이동합니다.\n" +
      "이동된 원본 9행 블록은 통합관리시트에서 삭제됩니다.\n\n실행할까요?",
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  try {
    var result = moveCompletedProjectsToCompletedSheet_();
    var msg = [
      "✅ 완료 프로젝트 이동 완료",
      "이동: " + result.moved + "건",
      "스킵: " + result.skipped + "건",
      "오류: " + result.errors + "건"
    ];

    if (result.copiedButNotDeleted > 0) {
      msg.push("");
      msg.push("⚠️ 복사는 됐지만 원본 삭제가 실패한 건: " + result.copiedButNotDeleted + "건");
    }
    if (result.movedNames && result.movedNames.length > 0) {
      msg.push("");
      msg.push("이동한 프로젝트:");
      result.movedNames.slice(0, 12).forEach(function(name) {
        msg.push("- " + name);
      });
      if (result.movedNames.length > 12) {
        msg.push("- 외 " + (result.movedNames.length - 12) + "건");
      }
    }
    if (result.errorMessages && result.errorMessages.length > 0) {
      msg.push("");
      msg.push("오류:");
      result.errorMessages.slice(0, 5).forEach(function(errorMsg) {
        msg.push("- " + errorMsg);
      });
    }

    ui.alert(msg.join("\n"));
  } catch (e) {
    ui.alert("❌ 완료 프로젝트 이동 실패\n" + (e && e.message ? e.message : e));
  }
}

function moveCompletedProjectsToCompletedSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = getMainSheet_();
  var targetSheet = getOrCreateCompletedSheet_(ss, sourceSheet);

  if (sourceSheet.getSheetId() === targetSheet.getSheetId()) {
    throw new Error("원본 시트와 완료 탭이 같습니다: " + sourceSheet.getName());
  }

  var blockHeight = getBlockHeight_(sourceSheet);
  var startRow = getStartRow_();
  var lastRow = sourceSheet.getLastRow();
  var lastCol = sourceSheet.getLastColumn();
  var result = {
    moved: 0,
    skipped: 0,
    errors: 0,
    copiedButNotDeleted: 0,
    movedNames: [],
    errorMessages: []
  };

  if (lastRow < startRow) return result;

  ensureSheetSizeForRange_(targetSheet, startRow, lastCol);
  copyColumnWidths_(sourceSheet, targetSheet, lastCol);

  var stopCtl = makeStopController_();
  var candidates = [];

  for (var r = startRow; r <= lastRow; r += blockHeight) {
    if (stopCtl.check(sourceSheet, r)) break;

    var projectName = sourceSheet.getRange(r + CONFIG.POS_NAME.row, CONFIG.POS_NAME.col).getDisplayValue();
    if (!isValidName(projectName)) continue;

    var status = sourceSheet.getRange(r, 7).getDisplayValue();
    status = String(status || "").trim();
    if (status !== "완료") {
      result.skipped++;
      continue;
    }

    candidates.push({
      row: r,
      projectName: String(projectName || "").trim()
    });
  }

  if (candidates.length === 0) return result;

  var nextTargetRow = getNextCompletedAppendRow_(targetSheet, blockHeight);
  var copied = [];

  for (var i = 0; i < candidates.length; i++) {
    var item = candidates[i];
    try {
      ensureSheetSizeForRange_(targetSheet, nextTargetRow + blockHeight - 1, lastCol);
      copyProjectBlock_(sourceSheet, item.row, targetSheet, nextTargetRow, blockHeight, lastCol);
      copied.push({
        sourceRow: item.row,
        targetRow: nextTargetRow,
        projectName: item.projectName
      });
      nextTargetRow += blockHeight;
    } catch (e) {
      result.errors++;
      result.errorMessages.push("Row " + item.row + " 복사 실패: " + (e && e.message ? e.message : e));
    }
  }

  copied.sort(function(a, b) {
    return b.sourceRow - a.sourceRow;
  });

  for (var j = 0; j < copied.length; j++) {
    var copiedItem = copied[j];
    try {
      sourceSheet.deleteRows(copiedItem.sourceRow, blockHeight);
      result.moved++;
      result.movedNames.unshift(copiedItem.projectName);
    } catch (deleteError) {
      result.errors++;
      result.copiedButNotDeleted++;
      result.errorMessages.push(
        "Row " + copiedItem.sourceRow + " 원본 삭제 실패: " +
        (deleteError && deleteError.message ? deleteError.message : deleteError)
      );
    }
  }

  return result;
}

function getCompletedSheetName_() {
  return (CONFIG && CONFIG.COMPLETED_SHEET_NAME) ? CONFIG.COMPLETED_SHEET_NAME : "완료";
}

function getOrCreateCompletedSheet_(ss, sourceSheet) {
  var completedSheetName = getCompletedSheetName_();
  var sheet = ss.getSheetByName(completedSheetName);
  if (sheet) return sheet;

  sheet = ss.insertSheet(completedSheetName);
  var startRow = getStartRow_();
  var headerRows = startRow - 1;
  var lastCol = sourceSheet.getLastColumn();

  ensureSheetSizeForRange_(sheet, startRow, lastCol);
  if (headerRows > 0 && lastCol > 0) {
    sourceSheet.getRange(1, 1, headerRows, lastCol).copyTo(
      sheet.getRange(1, 1, headerRows, lastCol),
      { contentsOnly: false }
    );
    copyRowHeights_(sourceSheet, 1, sheet, 1, headerRows);
    copyColumnWidths_(sourceSheet, sheet, lastCol);
  }
  return sheet;
}

function getNextCompletedAppendRow_(sheet, blockHeight) {
  var startRow = getStartRow_();
  var lastRow = sheet.getLastRow();
  if (lastRow < startRow) return startRow;

  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var rowCount = lastRow - startRow + 1;
  var values = sheet.getRange(startRow, 1, rowCount, lastCol).getDisplayValues();
  var lastUsedBlockStart = null;

  for (var offset = 0; offset < values.length; offset += blockHeight) {
    var blockHasValue = false;
    var end = Math.min(offset + blockHeight, values.length);
    for (var r = offset; r < end; r++) {
      for (var c = 0; c < values[r].length; c++) {
        if (String(values[r][c] || "").trim() !== "") {
          blockHasValue = true;
          break;
        }
      }
      if (blockHasValue) break;
    }
    if (blockHasValue) lastUsedBlockStart = startRow + offset;
  }

  return lastUsedBlockStart === null ? startRow : lastUsedBlockStart + blockHeight;
}

function copyProjectBlock_(sourceSheet, sourceRow, targetSheet, targetRow, blockHeight, lastCol) {
  var sourceRange = sourceSheet.getRange(sourceRow, 1, blockHeight, lastCol);
  var targetRange = targetSheet.getRange(targetRow, 1, blockHeight, lastCol);
  sourceRange.copyTo(targetRange, { contentsOnly: false });
  copyRowHeights_(sourceSheet, sourceRow, targetSheet, targetRow, blockHeight);
}

function ensureSheetSizeForRange_(sheet, requiredLastRow, requiredLastCol) {
  if (sheet.getMaxRows() < requiredLastRow) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredLastRow - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < requiredLastCol) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredLastCol - sheet.getMaxColumns());
  }
}

function copyRowHeights_(sourceSheet, sourceStartRow, targetSheet, targetStartRow, rowCount) {
  for (var i = 0; i < rowCount; i++) {
    targetSheet.setRowHeight(targetStartRow + i, sourceSheet.getRowHeight(sourceStartRow + i));
  }
}

function copyColumnWidths_(sourceSheet, targetSheet, colCount) {
  for (var c = 1; c <= colCount; c++) {
    targetSheet.setColumnWidth(c, sourceSheet.getColumnWidth(c));
  }
}
