/**
 * 인테리어 DB 동기화 핵심 서비스
 */

function runInteriorDbSync() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = getUiIfAvailable_();
  var lock = LockService.getDocumentLock();

  if (!lock.tryLock(30000)) {
    ss.toast('이미 동기화가 실행 중입니다. 잠시 후 다시 시도해주세요.', '🛋️ 인테리어 관리', 5);
    alertIfPossible_(ui, '이미 동기화가 실행 중입니다. 잠시 후 다시 시도해주세요.');
    return;
  }

  try {
    var sourceSheet = getSheetByAliases_(ss, INTERIOR_SYNC_CONFIG.SOURCE_SHEET_ALIASES);
    var clientsSheet = getOrCreateTargetSheet_(ss, INTERIOR_SYNC_CONFIG.TARGET_CLIENTS_ALIASES, INTERIOR_SYNC_CONFIG.TARGET_CLIENTS, INTERIOR_SYNC_CONFIG.TARGET_HEADERS.clients);
    var projectsSheet = getOrCreateTargetSheet_(ss, INTERIOR_SYNC_CONFIG.TARGET_PROJECTS_ALIASES, INTERIOR_SYNC_CONFIG.TARGET_PROJECTS, INTERIOR_SYNC_CONFIG.TARGET_HEADERS.projects);
    var milestonesSheet = getOrCreateTargetSheet_(ss, INTERIOR_SYNC_CONFIG.TARGET_MILESTONES_ALIASES, INTERIOR_SYNC_CONFIG.TARGET_MILESTONES, INTERIOR_SYNC_CONFIG.TARGET_HEADERS.milestones);

    if (!sourceSheet || !clientsSheet || !projectsSheet || !milestonesSheet) {
      var missing = [];
      if (!sourceSheet) missing.push('통합관리시트');
      if (!clientsSheet) missing.push('clients');
      if (!projectsSheet) missing.push('projects');
      if (!milestonesSheet) missing.push('milestones');
      throw new Error('필수 시트를 찾을 수 없습니다. 누락: ' + missing.join('/'));
    }

    var anchors = collectAnchorRows_(sourceSheet);
    if (anchors.length === 0) {
      ss.toast('동기화할 프로젝트 코드가 없습니다.', '🛋️ 인테리어 관리', 5);
      return;
    }

    var clientsRows = [];
    var projectsRows = [];
    var milestonesRows = [];
    var projectCodesToRefresh = {};

    anchors.forEach(function(anchorRow, idx) {
      var nextAnchorRow = (idx + 1 < anchors.length) ? anchors[idx + 1] : null;
      var record = buildRecordFromAnchor_(sourceSheet, anchorRow, nextAnchorRow);
      if (!record || !record.projectCode) return;

      clientsRows.push([record.clientId, record.clientName, record.phone]);
      projectsRows.push([
        record.projectCode,
        record.clientId,
        record.clientName,
        record.projectType,
        record.contractDate,
        record.balanceDate,
        record.address,
        record.memo,
        record.addressLink,
        record.folderLink,
        record.beforePhotoLink,
        record.constructionPhotoLink,
        record.afterPhotoLink,
        record.aviLink,
        record.blogLink,
        record.viewerLink,
        record.editLink,
        record.sheetLink
      ]);

      projectCodesToRefresh[record.projectCode] = true;
      Array.prototype.push.apply(milestonesRows, record.milestones);
    });

    var clientsResult = upsertByKey_(clientsSheet, clientsRows, 1);
    var projectsResult = upsertByKey_(projectsSheet, projectsRows, 1);
    replaceMilestonesByProjectCodes_(milestonesSheet, Object.keys(projectCodesToRefresh), milestonesRows);

    ss.toast(
      '동기화 완료 - clients:' + clientsResult.applied + ' / projects:' + projectsResult.applied + ' / milestones:' + milestonesRows.length,
      '🛋️ 인테리어 관리',
      5
    );
  } catch (err) {
    alertIfPossible_(ui, '동기화 중 오류가 발생했습니다.\n' + err.message);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function upsertByKey_(targetSheet, rows, keyColIndex1Based) {
  if (!rows || rows.length === 0) {
    return { applied: 0, inserted: 0, updated: 0, skippedEmptyKey: 0 };
  }

  var dataStartRow = 2;
  var lastRow = targetSheet.getLastRow();
  var keyToRowMap = {};

  if (lastRow >= dataStartRow) {
    var existingValues = targetSheet.getRange(dataStartRow, 1, lastRow - 1, targetSheet.getLastColumn()).getValues();
    for (var i = 0; i < existingValues.length; i++) {
      var key = (existingValues[i][keyColIndex1Based - 1] || '').toString().trim();
      if (key) keyToRowMap[key] = dataStartRow + i;
    }
  }

  var appendRows = [];
  var updatedCount = 0;
  var skippedEmptyKeyCount = 0;

  rows.forEach(function(row) {
    var key = (row[keyColIndex1Based - 1] || '').toString().trim();
    if (!key) {
      skippedEmptyKeyCount++;
      return;
    }

    if (keyToRowMap[key]) {
      targetSheet.getRange(keyToRowMap[key], 1, 1, row.length).setValues([row]);
      updatedCount++;
    } else {
      appendRows.push(row);
    }
  });

  if (appendRows.length > 0) {
    var appendStart = targetSheet.getLastRow() + 1;
    targetSheet.getRange(appendStart, 1, appendRows.length, appendRows[0].length).setValues(appendRows);
  }

  return {
    applied: updatedCount + appendRows.length,
    inserted: appendRows.length,
    updated: updatedCount,
    skippedEmptyKey: skippedEmptyKeyCount
  };
}

function replaceMilestonesByProjectCodes_(milestonesSheet, projectCodes, newRows) {
  var dataStartRow = 2;
  var lastRow = milestonesSheet.getLastRow();
  var codeMap = {};

  projectCodes.forEach(function(code) {
    if (code) codeMap[code] = true;
  });

  var keepRows = [];
  if (lastRow >= dataStartRow) {
    var existing = milestonesSheet.getRange(dataStartRow, 1, lastRow - 1, milestonesSheet.getLastColumn()).getValues();
    for (var i = 0; i < existing.length; i++) {
      var code = (existing[i][0] || '').toString().trim();
      if (!codeMap[code]) keepRows.push(existing[i]);
    }
  }

  var finalRows = keepRows.concat(newRows || []);

  var maxCols = milestonesSheet.getLastColumn();
  if (lastRow >= dataStartRow) {
    milestonesSheet.getRange(dataStartRow, 1, lastRow - 1, maxCols).clearContent();
  }

  if (finalRows.length > 0) {
    milestonesSheet.getRange(dataStartRow, 1, finalRows.length, finalRows[0].length).setValues(finalRows);
  }
}
