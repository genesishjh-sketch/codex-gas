/** FolderFileManager.gs
 *
 * 폴더 생성:
 * - 블록 기준 r+1~r+4 (R열) 텍스트를 폴더명으로 사용
 * - 상위 폴더에 생성, S열에 링크 삽입
 * - 링크 있으면 유지, 같은 이름 폴더는 재사용
 * - 완료/취소 블록 스킵
 * - 연속 빈 블록 N개면 중단
 */

function createFoldersBatch(isSilent, force) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getMainSheet_();
  var blockHeight = getBlockHeight_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.START_ROW) return { summary: "No data", successList: [], failedList: [] };

  var stopCtl = makeStopController_();
  var parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();

  var processedCount = 0;
  var skipCount = 0;
  var successList = [];
  var failedList = [];

  var labelCol = (CONFIG && CONFIG.POS_FOLDER_LABEL_COL) || 18; // R
  var urlCol = (CONFIG && (CONFIG.POS_FOLDER_URL_COL || CONFIG.DRIVE_MARK_COL)) || 19; // S

  for (var r = CONFIG.START_ROW; r <= lastRow; r += blockHeight) {
    if (stopCtl.check(sheet, r)) break;

    try {
      var nameVal = sheet.getRange(r, CONFIG.POS_NAME.col).getDisplayValue();
      if (!isValidName(nameVal)) continue;

      if (isClosedBlock_(sheet, r)) { skipCount++; continue; }

      var projectFolder = getOrCreateProjectFolder_(parentFolder, sheet, r);

      // rows: r+1..r+4 (R5-8, R13-16, ...)
      for (var i = 1; i <= 4; i++) {
        var row = r + i;
        if (row > lastRow) break;

        var labelCell = sheet.getRange(row, labelCol);
        var label = (labelCell.getDisplayValue() || "").toString().trim();
        if (!label) continue;

        var urlCell = sheet.getRange(row, urlCol);
        var url = (urlCell.getDisplayValue() || "").toString().trim();
        if (!url) {
          url = getUrlFromCell_(urlCell);
        }

        if (!force && url && url.indexOf("drive.google.com") >= 0) {
          skipCount++;
          continue;
        }

        var folders = projectFolder.getFoldersByName(label);
        var folder = folders.hasNext() ? folders.next() : projectFolder.createFolder(label);
        urlCell.setValue(folder.getUrl());
        processedCount++;
        successList.push(label);
      }

    } catch (e) {
      failedList.push("Row " + r + ": " + e.message);
    }
  }

  var summary = "Processed " + processedCount + "/ Skipped " + skipCount;
  if (!isSilent) SpreadsheetApp.getUi().alert("Folder create done\n" + summary);

  return { summary: summary, successList: successList, failedList: failedList };
}

function getOrCreateProjectFolder_(parentFolder, sheet, blockStartRow) {
  var nameCell = sheet.getRange(blockStartRow + CONFIG.POS_NAME.row, 3);

  var nameVal = (nameCell.getDisplayValue() || "").toString().trim();

  var dateCell = sheet.getRange(blockStartRow + 4, 4);
  var dateVal = dateCell.getValue();
  var dateStr = formatProjectDate_(dateVal, dateCell);

  var projectName = (dateStr ? dateStr + " " : "") + nameVal;
  projectName = projectName.trim() || "프로젝트";

  var existing = parentFolder.getFoldersByName(projectName);
  return existing.hasNext() ? existing.next() : parentFolder.createFolder(projectName);
}

function formatProjectDate_(value, cell) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyMMdd");
  }
  var display = (cell && cell.getDisplayValue) ? cell.getDisplayValue() : "";
  display = (display || "").toString().trim();
  if (!display) return "";

  var digits = display.replace(/[^\d]/g, "");
  if (digits.length === 8) return digits.slice(2);
  if (digits.length === 6) return digits;
  return display;
}
