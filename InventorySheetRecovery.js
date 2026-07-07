/** InventorySheetRecovery.gs
 * 현재 선택된 프로젝트의 물품리스트 구글시트만 복구한다.
 */

function runRecoverInventorySheetForCurrentProject() {
  var ui = SpreadsheetApp.getUi();
  try {
    var result = recoverInventorySheetForCurrentProject_();
    var actionText = {
      EXISTS: "기존 물품리스트가 정상이라 새로 만들지 않았습니다.",
      RELINKED: "프로젝트 폴더에서 기존 물품리스트를 찾아 링크를 복구했습니다.",
      CREATED: "템플릿에서 새 물품리스트 구글시트를 만들었습니다."
    };

    ui.alert([
      "📄 물품리스트 복구 완료",
      "프로젝트: " + result.projectName,
      "시트: " + result.sheetName + " / Row " + result.blockStartRow,
      "결과: " + (actionText[result.action] || result.action),
      "링크: " + result.fileUrl
    ].join("\n"));
  } catch (e) {
    ui.alert("❌ 물품리스트 복구 실패\n" + (e && e.message ? e.message : e));
  }
}

function recoverInventorySheetForCurrentProject_() {
  var context = resolveCurrentProjectForInventoryRecovery_();
  var sheet = context.sheet;
  var blockStartRow = context.blockStartRow;
  var projectName = context.projectName;
  var fileCell = sheet.getRange(blockStartRow + CONFIG.POS_FILE.row, CONFIG.POS_FILE.col);
  var existingUrl = readUrlFromCell_(fileCell);
  var existingFile = getActiveSpreadsheetFileFromUrl_(existingUrl);

  if (existingFile) {
    setInventorySheetProjectName_(existingFile.getId(), sheet, blockStartRow);
    fileCell.setValue(existingFile.getUrl());
    return {
      action: "EXISTS",
      sheetName: sheet.getName(),
      blockStartRow: blockStartRow,
      projectName: projectName,
      fileUrl: existingFile.getUrl()
    };
  }

  var projectFolder = resolveProjectFolderForInventoryRecovery_(sheet, blockStartRow);
  var fileName = buildProjectFileName_(sheet, blockStartRow);
  var relinkFile = findActiveSpreadsheetFileByName_(projectFolder, fileName);
  if (relinkFile) {
    setInventorySheetProjectName_(relinkFile.getId(), sheet, blockStartRow);
    fileCell.setValue(relinkFile.getUrl());
    return {
      action: "RELINKED",
      sheetName: sheet.getName(),
      blockStartRow: blockStartRow,
      projectName: projectName,
      fileUrl: relinkFile.getUrl()
    };
  }

  var templateId = getInventoryTemplateIdForRecovery_(sheet);
  if (!templateId) throw new Error("G1에서 물품리스트 템플릿 URL을 찾을 수 없습니다.");

  var newFile = copyTemplateFile_(templateId, fileName, projectFolder);
  setInventorySheetProjectName_(newFile.getId(), sheet, blockStartRow);
  fileCell.setValue(newFile.getUrl());

  return {
    action: "CREATED",
    sheetName: sheet.getName(),
    blockStartRow: blockStartRow,
    projectName: projectName,
    fileUrl: newFile.getUrl()
  };
}

function resolveCurrentProjectForInventoryRecovery_() {
  var activeRange = SpreadsheetApp.getActiveRange();
  if (!activeRange) throw new Error("현재 선택된 셀이 없습니다. 복구할 프로젝트 안의 셀을 선택해주세요.");

  var sheet = activeRange.getSheet();
  if (!isInventoryRecoverySheet_(sheet)) {
    throw new Error("통합관리시트 또는 완료 탭에서 복구할 프로젝트 안의 셀을 선택해주세요.");
  }

  var blockHeight = getBlockHeight_(sheet);
  var blockStartRow = getBlockStartRow_(activeRange.getRow(), blockHeight);
  if (!blockStartRow) throw new Error("프로젝트 영역 안의 셀을 선택해주세요.");

  var projectName = sheet.getRange(blockStartRow + CONFIG.POS_NAME.row, CONFIG.POS_NAME.col).getDisplayValue();
  projectName = String(projectName || "").trim();
  if (!isValidName(projectName)) {
    throw new Error("현재 선택 위치에서 유효한 프로젝트명을 찾지 못했습니다. Row " + blockStartRow + "을 확인해주세요.");
  }

  return {
    sheet: sheet,
    blockStartRow: blockStartRow,
    projectName: projectName
  };
}

function isInventoryRecoverySheet_(sheet) {
  if (!sheet) return false;
  var name = sheet.getName();
  var mainName = (CONFIG && CONFIG.SHEET_NAME) ? CONFIG.SHEET_NAME : "통합관리시트";
  var completedName = (CONFIG && CONFIG.COMPLETED_SHEET_NAME) ? CONFIG.COMPLETED_SHEET_NAME : "완료";
  return name === mainName || name === completedName;
}

function readUrlFromCell_(cell) {
  var linkedUrl = getUrlFromCell_(cell);
  if (linkedUrl) return linkedUrl;

  var value = (cell.getDisplayValue() || "").toString().trim();
  return value || "";
}

function getActiveSpreadsheetFileFromUrl_(url) {
  var id = extractUsableDriveId_(url);
  if (!id) return null;

  try {
    var file = DriveApp.getFileById(id);
    if (typeof file.isTrashed === "function" && file.isTrashed()) return null;
    if (!canOpenAsSpreadsheet_(file)) return null;
    return file;
  } catch (e) {
    return null;
  }
}

function extractUsableDriveId_(url) {
  var id = extractIdFromUrl(url);
  if (!id || id.indexOf("http") >= 0 || id.indexOf("/") >= 0) return "";
  return id;
}

function canOpenAsSpreadsheet_(file) {
  try {
    SpreadsheetApp.openById(file.getId());
    return true;
  } catch (e) {
    return false;
  }
}

function resolveProjectFolderForInventoryRecovery_(sheet, blockStartRow) {
  var urlCol = (CONFIG && (CONFIG.POS_FOLDER_URL_COL || CONFIG.DRIVE_MARK_COL)) || 19;
  var folderUrl = readUrlFromCell_(sheet.getRange(blockStartRow, urlCol));
  var folder = getActiveFolderFromUrl_(folderUrl);
  if (folder) return folder;

  folder = findProjectFolderByNameForRecovery_(sheet, blockStartRow);
  if (folder) return folder;

  throw new Error("프로젝트 폴더를 찾지 못했습니다. S열의 프로젝트 폴더 링크를 확인해주세요.");
}

function getActiveFolderFromUrl_(url) {
  var id = extractUsableDriveId_(url);
  if (!id) return null;

  try {
    var folder = DriveApp.getFolderById(id);
    if (typeof folder.isTrashed === "function" && folder.isTrashed()) return null;
    return folder;
  } catch (e) {
    return null;
  }
}

function findProjectFolderByNameForRecovery_(sheet, blockStartRow) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetParentFolder = DriveApp.getFileById(ss.getId()).getParents().next();
  var rootFolders = spreadsheetParentFolder.getFoldersByName("01 프로젝트관리");
  if (!rootFolders.hasNext()) return null;

  var projectFolderName = buildProjectFolderNameForRecovery_(sheet, blockStartRow);
  var folders = rootFolders.next().getFoldersByName(projectFolderName);
  while (folders.hasNext()) {
    var folder = folders.next();
    if (typeof folder.isTrashed !== "function" || !folder.isTrashed()) return folder;
  }
  return null;
}

function buildProjectFolderNameForRecovery_(sheet, blockStartRow) {
  var nameVal = sheet.getRange(blockStartRow + CONFIG.POS_NAME.row, CONFIG.POS_NAME.col).getDisplayValue();
  nameVal = String(nameVal || "").trim();

  var dateCell = sheet.getRange(blockStartRow + 5, 4);
  var dateStr = formatProjectDate_(dateCell.getValue(), dateCell);
  return ((dateStr ? dateStr + " " : "") + nameVal).trim() || "프로젝트";
}

function findActiveSpreadsheetFileByName_(folder, fileName) {
  var files = folder.getFilesByName(fileName);
  while (files.hasNext()) {
    var file = files.next();
    if (typeof file.isTrashed === "function" && file.isTrashed()) continue;
    if (canOpenAsSpreadsheet_(file)) return file;
  }
  return null;
}

function getInventoryTemplateIdForRecovery_(sheet) {
  var templateUrl = getTemplateUrl_(sheet);
  if (!templateUrl && sheet.getName() !== CONFIG.SHEET_NAME) {
    try {
      templateUrl = getTemplateUrl_(getMainSheet_());
    } catch (e) {}
  }
  return extractUsableDriveId_(templateUrl);
}

function setInventorySheetProjectName_(spreadsheetId, sourceSheet, blockStartRow) {
  var sourceText = sourceSheet.getRange(blockStartRow, CONFIG.POS_NAME.col).getDisplayValue();
  var fileSs = SpreadsheetApp.openById(spreadsheetId);
  var targetSheet = fileSs.getSheets()[0];
  targetSheet.getRange("B3").setValue(sourceText);
}
