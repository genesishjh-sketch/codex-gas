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

    var recordByProjectCode = {};

    anchors.forEach(function(anchorRow, idx) {
      var nextAnchorRow = (idx + 1 < anchors.length) ? anchors[idx + 1] : null;
      var record = buildRecordFromAnchor_(sourceSheet, anchorRow, nextAnchorRow);
      if (!record || !record.projectCode) return;
      // 같은 프로젝트 코드가 통합관리시트에 중복 존재하면 마지막(아래) 블록을 우선 사용
      // (기존엔 동일 코드의 마일스톤이 누적되어 DB에서 일정이 섞여 보일 수 있음)
      recordByProjectCode[record.projectCode] = record;
    });

    var clientsRows = [];
    var projectsRows = [];
    var milestonesRows = [];
    var projectCodesToRefresh = Object.keys(recordByProjectCode);

    projectCodesToRefresh.forEach(function(projectCode) {
      var record = recordByProjectCode[projectCode];
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

      Array.prototype.push.apply(milestonesRows, record.milestones || []);
    });

    var clientsResult = upsertByKey_(clientsSheet, clientsRows, 1);
    var projectsResult = upsertByKey_(projectsSheet, projectsRows, 1);
    replaceMilestonesByProjectCodes_(milestonesSheet, projectCodesToRefresh, milestonesRows);

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
  var baseMilestoneColCount = INTERIOR_SYNC_CONFIG.TARGET_HEADERS.milestones.length;
  var maxCols = milestonesSheet.getLastColumn();

  projectCodes.forEach(function(code) {
    if (code) codeMap[code] = true;
  });

  var keepRows = [];
  var preservedMetaByKey = {};
  if (lastRow >= dataStartRow) {
    var existing = milestonesSheet.getRange(dataStartRow, 1, lastRow - 1, maxCols).getValues();
    for (var i = 0; i < existing.length; i++) {
      var code = (existing[i][0] || '').toString().trim();
      if (!codeMap[code]) {
        keepRows.push(existing[i]);
        continue;
      }

      var meta = extractMilestoneMeta_(existing[i], baseMilestoneColCount);
      if (!meta.hasMeta) continue;

      var key = makeMilestoneIdentityKey_(existing[i]);
      if (!key) continue;
      preservedMetaByKey[key] = meta;
    }
  }

  var restoredRows = (newRows || []).map(function(row) {
    var rowCopy = row.slice();
    var key = makeMilestoneIdentityKey_(rowCopy);
    if (!key || !preservedMetaByKey[key]) return rowCopy;
    return applyMilestoneMeta_(rowCopy, preservedMetaByKey[key], baseMilestoneColCount);
  });

  var finalRows = keepRows.concat(restoredRows);
  if (lastRow >= dataStartRow) {
    milestonesSheet.getRange(dataStartRow, 1, lastRow - 1, maxCols).clearContent();
  }

  if (finalRows.length > 0) {
    milestonesSheet.getRange(dataStartRow, 1, finalRows.length, finalRows[0].length).setValues(finalRows);
  }
}

function makeMilestoneIdentityKey_(row) {
  if (!row || row.length < 4) return '';
  var projectCode = normalizeIdentityPart_(row[0]);
  var section = normalizeIdentityPart_(row[1]);
  var stepName = normalizeIdentityPart_(row[2]);
  var planDate = normalizeDateIdentityPart_(row[3]);

  if (!projectCode || !planDate) return '';
  return [projectCode, section, stepName, planDate].join('||');
}

function normalizeIdentityPart_(value) {
  return (value || '').toString().trim().toLowerCase();
}

function normalizeDateIdentityPart_(value) {
  if (!value) return '';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd');
  }

  return (value || '').toString().trim();
}

function extractMilestoneMeta_(row, baseColCount) {
  var normalizedBase = Math.max(0, Number(baseColCount) || 0);
  var meta = {
    hasMeta: false,
    todoist_task_id: '',
    sync_status: '',
    last_synced_at: '',
    last_error: '',
    process_mark: ''
  };

  if (!row || row.length <= normalizedBase) return meta;

  function read(offset) {
    return row[normalizedBase + offset];
  }

  meta.todoist_task_id = read(0) || '';
  meta.sync_status = read(1) || '';
  meta.last_synced_at = read(2) || '';
  meta.last_error = read(3) || '';
  meta.process_mark = read(5) || '';
  meta.hasMeta = !!(meta.todoist_task_id || meta.sync_status || meta.last_synced_at || meta.last_error || meta.process_mark);
  return meta;
}

function applyMilestoneMeta_(row, meta, baseColCount) {
  var output = row.slice();
  var normalizedBase = Math.max(0, Number(baseColCount) || 0);

  while (output.length < normalizedBase + 6) {
    output.push('');
  }

  output[normalizedBase] = meta.todoist_task_id || '';
  output[normalizedBase + 1] = meta.sync_status || '';
  output[normalizedBase + 2] = meta.last_synced_at || '';
  output[normalizedBase + 3] = meta.last_error || '';
  output[normalizedBase + 5] = meta.process_mark || '';
  return output;
}
