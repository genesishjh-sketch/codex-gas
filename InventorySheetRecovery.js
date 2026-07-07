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

    var msg = [
      "📄 물품리스트 복구 완료",
      "프로젝트: " + result.projectName,
      "시트: " + result.sheetName + " / Row " + result.blockStartRow,
      "결과: " + (actionText[result.action] || result.action),
      "링크: " + result.fileUrl
    ];

    if (result.postProcess && result.postProcess.duplicateProtectionsRemoved > 0) {
      msg.push("중복 보호범위 정리: " + result.postProcess.duplicateProtectionsRemoved + "개");
    }
    if (result.postProcess && result.postProcess.warnings && result.postProcess.warnings.length > 0) {
      msg.push("");
      msg.push("경고:");
      result.postProcess.warnings.slice(0, 5).forEach(function(warning) {
        msg.push("- " + warning);
      });
    }

    ui.alert(msg.join("\n"));
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
    var existingFinalize = finalizeInventorySheetFile_(existingFile.getId(), sheet, blockStartRow);
    fileCell.setValue(existingFile.getUrl());
    return {
      action: "EXISTS",
      sheetName: sheet.getName(),
      blockStartRow: blockStartRow,
      projectName: projectName,
      fileUrl: existingFile.getUrl(),
      postProcess: existingFinalize
    };
  }

  var projectFolder = resolveProjectFolderForInventoryRecovery_(sheet, blockStartRow);
  var fileName = buildProjectFileName_(sheet, blockStartRow);
  var relinkFile = findActiveSpreadsheetFileByName_(projectFolder, fileName);
  if (relinkFile) {
    var relinkFinalize = finalizeInventorySheetFile_(relinkFile.getId(), sheet, blockStartRow);
    fileCell.setValue(relinkFile.getUrl());
    return {
      action: "RELINKED",
      sheetName: sheet.getName(),
      blockStartRow: blockStartRow,
      projectName: projectName,
      fileUrl: relinkFile.getUrl(),
      postProcess: relinkFinalize
    };
  }

  var templateId = getInventoryTemplateIdForRecovery_(sheet);
  if (!templateId) throw new Error("G1에서 물품리스트 템플릿 URL을 찾을 수 없습니다.");

  var newFile = copyTemplateFile_(templateId, fileName, projectFolder);
  var newFinalize = finalizeInventorySheetFile_(newFile.getId(), sheet, blockStartRow);
  fileCell.setValue(newFile.getUrl());

  return {
    action: "CREATED",
    sheetName: sheet.getName(),
    blockStartRow: blockStartRow,
    projectName: projectName,
    fileUrl: newFile.getUrl(),
    postProcess: newFinalize
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

function finalizeInventorySheetFile_(spreadsheetId, sourceSheet, blockStartRow) {
  var result = {
    b3Updated: false,
    duplicateProtectionsRemoved: 0,
    warnings: []
  };

  try {
    setInventorySheetProjectName_(spreadsheetId, sourceSheet, blockStartRow);
    result.b3Updated = true;
  } catch (e) {
    result.warnings.push("B3 프로젝트명 입력 실패: " + (e && e.message ? e.message : e));
  }

  try {
    var protectionResult = removeExactDuplicateRangeProtections_(spreadsheetId);
    result.duplicateProtectionsRemoved = protectionResult.removed;
    result.warnings = result.warnings.concat(protectionResult.warnings);
  } catch (cleanupError) {
    result.warnings.push("중복 보호범위 정리 실패: " + (cleanupError && cleanupError.message ? cleanupError.message : cleanupError));
  }

  return result;
}

function removeExactDuplicateRangeProtections_(spreadsheetId) {
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var protections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  var seen = {};
  var removed = 0;
  var warnings = [];

  for (var i = 0; i < protections.length; i++) {
    var protection = protections[i];
    var key = getRangeProtectionFingerprint_(protection);
    if (!key) continue;

    if (!seen[key]) {
      seen[key] = true;
      continue;
    }

    try {
      if (typeof protection.canEdit === "function" && !protection.canEdit()) {
        warnings.push("삭제 권한 없는 중복 보호범위 유지: " + describeRangeProtection_(protection));
        continue;
      }
      protection.remove();
      removed++;
    } catch (e) {
      warnings.push("중복 보호범위 삭제 실패: " + describeRangeProtection_(protection) + " - " + (e && e.message ? e.message : e));
    }
  }

  return { removed: removed, warnings: warnings };
}

function getRangeProtectionFingerprint_(protection) {
  try {
    var range = protection.getRange();
    if (!range) return "";

    var sheet = range.getSheet();
    var parts = [
      sheet.getSheetId(),
      range.getA1Notation(),
      String(protection.getDescription() || ""),
      protection.isWarningOnly() ? "warning" : "restricted"
    ];

    if (!protection.isWarningOnly()) {
      var editorsFingerprint = getProtectionEditorsFingerprint_(protection);
      var domainFingerprint = getProtectionDomainFingerprint_(protection);
      if (editorsFingerprint === null || domainFingerprint === null) return "";
      parts.push(editorsFingerprint);
      parts.push(domainFingerprint);
    }

    return parts.join("\u001f");
  } catch (e) {
    return "";
  }
}

function getProtectionEditorsFingerprint_(protection) {
  try {
    var editors = protection.getEditors() || [];
    var emails = editors.map(function(editor) {
      return (editor && editor.getEmail) ? editor.getEmail() : String(editor || "");
    }).filter(function(email) {
      return !!email;
    }).sort();
    return emails.join(",");
  } catch (e) {
    return null;
  }
}

function getProtectionDomainFingerprint_(protection) {
  try {
    return protection.canDomainEdit() ? "domain_edit" : "domain_locked";
  } catch (e) {
    return null;
  }
}

function describeRangeProtection_(protection) {
  try {
    var range = protection.getRange();
    return range.getSheet().getName() + "!" + range.getA1Notation();
  } catch (e) {
    return "(알 수 없는 범위)";
  }
}
