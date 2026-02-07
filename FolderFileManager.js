/** FolderFileManager.gs
 *
 * 폴더 생성:
 * - 블록 기준 r+1~r+4 (R열) 텍스트를 폴더명으로 사용
 * - 스프레드시트 부모 폴더의 "01 프로젝트관리" 하위에 생성, S열에 링크 삽입
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
  var spreadsheetParentFolder = DriveApp.getFileById(ss.getId()).getParents().next();
  var parentFolder = getOrCreateProjectRootFolder_(spreadsheetParentFolder);
  var templateUrl = getTemplateUrl_(sheet);
  var templateId = templateUrl ? extractIdFromUrl(templateUrl) : "";

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
      var projectUrlCell = sheet.getRange(r, urlCol);
      var projectUrl = (projectUrlCell.getDisplayValue() || "").toString().trim();
      if (!projectUrl) {
        projectUrl = getUrlFromCell_(projectUrlCell);
      }
      if (force || !projectUrl || projectUrl.indexOf("drive.google.com") < 0) {
        projectUrlCell.setValue(projectFolder.getUrl());
      }

      var fileCell = sheet.getRange(r + CONFIG.POS_FILE.row, CONFIG.POS_FILE.col);
      var fileUrl = (fileCell.getDisplayValue() || "").toString().trim();
      if (!fileUrl) {
        fileUrl = getUrlFromCell_(fileCell);
      }
      var fileName = buildProjectFileName_(sheet, r);
      if (force || !fileUrl) {
        var file = findFileByName_(projectFolder, fileName);
        if (!file) {
          if (!templateId) throw new Error("물품리스트 템플릿 URL을 찾을 수 없습니다.");
          file = copyTemplateFile_(templateId, fileName, projectFolder);
          setImportRangeFormula_(file.getId(), r, blockHeight);
        } else if (force) {
          setImportRangeFormula_(file.getId(), r, blockHeight);
        }
        fileCell.setValue(file.getUrl());
      }

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

        var folder = getOrCreateSubFolder_(projectFolder, label);
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

  var dateCell = sheet.getRange(blockStartRow + 5, 4);
  var dateVal = dateCell.getValue();
  var dateStr = formatProjectDate_(dateVal, dateCell);

  var projectName = (dateStr ? dateStr + " " : "") + nameVal;
  projectName = projectName.trim() || "프로젝트";

  var existing = parentFolder.getFoldersByName(projectName);
  return existing.hasNext() ? existing.next() : parentFolder.createFolder(projectName);
}

function getOrCreateProjectRootFolder_(spreadsheetParentFolder) {
  var rootName = "01 프로젝트관리";
  var existing = spreadsheetParentFolder.getFoldersByName(rootName);
  return existing.hasNext() ? existing.next() : spreadsheetParentFolder.createFolder(rootName);
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

function getTemplateUrl_(sheet) {
  var templateCell = sheet.getRange((CONFIG && CONFIG.CELL_TEMPLATE_ORIGIN) || "G1");
  return getUrlFromCell_(templateCell) || templateCell.getDisplayValue();
}

function buildProjectFileName_(sheet, blockStartRow) {
  var nameCell = sheet.getRange(blockStartRow + CONFIG.POS_NAME.row, CONFIG.POS_NAME.col);
  var nameVal = (nameCell.getDisplayValue() || "").toString().trim();
  return (nameVal ? nameVal + " " : "") + "물품리스트";
}

function findFileByName_(folder, fileName) {
  var files = folder.getFilesByName(fileName);
  return files.hasNext() ? files.next() : null;
}

function copyTemplateFile_(templateId, fileName, parentFolder) {
  var templateFile = DriveApp.getFileById(templateId);
  return templateFile.makeCopy(fileName, parentFolder);
}

function setImportRangeFormula_(spreadsheetId, blockStartRow, blockHeight) {
  var sourceSpreadsheetId = "1GgdT1H-IEWWJDpuD14IURuZZWL9y_1smM9eUSXQMPoo";
  var sourceSheetName = "통합관리시트";
  var sourceSs = SpreadsheetApp.openById(sourceSpreadsheetId);
  var sourceSheet = sourceSs.getSheetByName(sourceSheetName);
  var sourceText = sourceSheet.getRange(blockStartRow, 3).getDisplayValue();
  var fileSs = SpreadsheetApp.openById(spreadsheetId);
  var targetSheet = fileSs.getSheets()[0];
  targetSheet.getRange("B3").setValue(sourceText);
}

function getOrCreateSubFolder_(projectFolder, label) {
  var folders = projectFolder.getFoldersByName(label);
  return folders.hasNext() ? folders.next() : projectFolder.createFolder(label);
}
