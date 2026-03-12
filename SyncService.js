/**
 * 인테리어 DB 동기화 핵심 서비스
 */

function runInteriorDbSync() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allAnchors = collectAnchorRowsForSync_(ss);
  syncInteriorDbByAnchors_(ss, allAnchors, {
    emptyMessage: '동기화할 프로젝트 코드가 없습니다.',
    donePrefix: '동기화 완료',
    toastTitle: '🛋️ 인테리어 관리',
    notifyUiOnError: true,
    verbose: true
  });
}

function runInteriorDbSyncRealtimeByEdit(e) {
  try {
    var anchors = collectEditedAnchorRows_(e);
    if (anchors.length === 0) return;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    syncInteriorDbByAnchors_(ss, anchors, {
      emptyMessage: '',
      donePrefix: '실시간 부분 동기화 완료',
      toastTitle: '🛋️ 인테리어 관리',
      notifyUiOnError: false,
      verbose: false
    });

    if (typeof runTodoistPendingQueueSync === 'function') {
      runTodoistPendingQueueSync();
    }
    if (typeof scheduleTodoistPendingQueueSyncFallback_ === 'function') {
      scheduleTodoistPendingQueueSyncFallback_();
    }
  } catch (err) {
    console.error('실시간 부분 동기화 실패: ' + (err && err.message ? err.message : err));
    throw err;
  }
}

function syncInteriorDbByAnchors_(ss, anchors, options) {
  options = options || {};
  var ui = getUiIfAvailable_();
  var lock = LockService.getDocumentLock();
  var normalizedAnchors = dedupeAnchorRows_(anchors);

  if (!lock.tryLock(30000)) {
    if (options.verbose !== false) {
      ss.toast('이미 동기화가 실행 중입니다. 잠시 후 다시 시도해주세요.', options.toastTitle || '🛋️ 인테리어 관리', 5);
      alertIfPossible_(ui, '이미 동기화가 실행 중입니다. 잠시 후 다시 시도해주세요.');
    }
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

    ensureSourceUidHeader_(sourceSheet);

    if (normalizedAnchors.length === 0) {
      if (options.emptyMessage) {
        ss.toast(options.emptyMessage, options.toastTitle || '🛋️ 인테리어 관리', 5);
      }
      return;
    }

    var recordByProjectCode = {};

    normalizedAnchors.forEach(function(anchorRow, idx) {
      var nextAnchorRow = (idx + 1 < normalizedAnchors.length) ? normalizedAnchors[idx + 1] : null;
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

    var donePrefix = options.donePrefix || '동기화 완료';
    if (options.verbose !== false) {
      ss.toast(
        donePrefix + ' - clients:' + clientsResult.applied + ' / projects:' + projectsResult.applied + ' / milestones:' + milestonesRows.length,
        options.toastTitle || '🛋️ 인테리어 관리',
        5
      );
    }
  } catch (err) {
    if (options.notifyUiOnError !== false) {
      alertIfPossible_(ui, '동기화 중 오류가 발생했습니다.\n' + err.message);
    }
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function collectAnchorRowsForSync_(ss) {
  var sourceSheet = getSheetByAliases_(ss, INTERIOR_SYNC_CONFIG.SOURCE_SHEET_ALIASES);
  if (!sourceSheet) return [];
  return collectAnchorRows_(sourceSheet);
}

function collectEditedAnchorRows_(e) {
  if (!e || !e.range) return [];

  var range = e.range;
  var sheet = range.getSheet();
  if (!sheet) return [];

  var sheetName = (sheet.getName() || '').toString().trim().toLowerCase();
  var aliases = INTERIOR_SYNC_CONFIG.SOURCE_SHEET_ALIASES || [INTERIOR_SYNC_CONFIG.SOURCE_SHEET];
  var isSourceSheet = aliases.some(function(alias) {
    return (alias || '').toString().trim().toLowerCase() === sheetName;
  });
  if (!isSourceSheet) return [];

  var startRow = range.getRow();
  var rowCount = range.getNumRows();
  if (!rowCount) return [];

  var blockStartRow = (typeof getStartRow_ === 'function') ? getStartRow_() : 4;
  var blockHeight = Math.max(1, (typeof getBlockHeight_ === 'function') ? getBlockHeight_(sheet) : 9);
  var projectCodeOffset = 7;
  var lastRow = sheet.getLastRow();
  var anchors = [];

  for (var row = startRow; row < startRow + rowCount; row++) {
    if (row < blockStartRow) continue;

    var blockIndex = Math.floor((row - blockStartRow) / blockHeight);
    var anchorRow = blockStartRow + (blockIndex * blockHeight) + projectCodeOffset;
    if (anchorRow < 1 || anchorRow > lastRow) continue;

    var projectCode = readCellDisplay_(sheet, anchorRow, 2);
    if (!projectCode || !isProjectCodeCandidate_(projectCode)) continue;
    anchors.push(anchorRow);
  }

  return dedupeAnchorRows_(anchors);
}

function dedupeAnchorRows_(anchors) {
  var map = {};
  var output = [];

  (anchors || []).forEach(function(anchorRow) {
    var numeric = Number(anchorRow);
    if (!numeric || numeric < 1) return;
    var normalized = Math.floor(numeric);
    if (map[normalized]) return;
    map[normalized] = true;
    output.push(normalized);
  });

  output.sort(function(a, b) { return a - b; });
  return output;
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
  var previousBaseByKey = {};
  var existingRowByKey = {};
  var removedMilestoneRows = [];
  if (lastRow >= dataStartRow) {
    var existing = milestonesSheet.getRange(dataStartRow, 1, lastRow - 1, maxCols).getValues();
    for (var i = 0; i < existing.length; i++) {
      var code = (existing[i][0] || '').toString().trim();
      if (!codeMap[code]) {
        keepRows.push(existing[i]);
        continue;
      }

      var key = makeMilestoneIdentityKey_(existing[i]);
      if (!key) continue;

      previousBaseByKey[key] = sliceMilestoneBaseRow_(existing[i], baseMilestoneColCount);

      var meta = extractMilestoneMeta_(existing[i], baseMilestoneColCount);
      preservedMetaByKey[key] = meta;
      existingRowByKey[key] = existing[i].slice();
    }
  }

  var newMilestoneKeySet = {};
  var restoredRows = (newRows || []).map(function(row) {
    var rowCopy = row.slice();
    var key = makeMilestoneIdentityKey_(rowCopy);
    if (!key) return rowCopy;

    newMilestoneKeySet[key] = true;

    var previousBase = previousBaseByKey[key];
    var currentBase = sliceMilestoneBaseRow_(rowCopy, baseMilestoneColCount);
    var isNew = !previousBase;
    var isChanged = !!previousBase && !isSameMilestoneBaseRow_(previousBase, currentBase);

    var meta = preservedMetaByKey[key] || {
      hasMeta: false,
      todoist_task_id: '',
      sync_status: '',
      last_synced_at: '',
      last_error: '',
      process_mark: ''
    };

    if (isNew) {
      meta.process_mark = '신규';
    } else if (isChanged) {
      meta.process_mark = '수정';
    }

    return applyMilestoneMeta_(rowCopy, meta, baseMilestoneColCount);
  });

  Object.keys(existingRowByKey).forEach(function(key) {
    if (newMilestoneKeySet[key]) return;
    removedMilestoneRows.push(existingRowByKey[key]);
  });

  cleanupRemovedMilestoneTodoistTasks_(removedMilestoneRows, baseMilestoneColCount);

  var finalRows = keepRows.concat(restoredRows);
  var writeWidth = Math.max(maxCols, baseMilestoneColCount + 5);

  if (finalRows.length > 0) {
    finalRows = normalizeRowsToWidth_(finalRows, writeWidth);
  }

  if (lastRow >= dataStartRow) {
    milestonesSheet.getRange(dataStartRow, 1, lastRow - 1, maxCols).clearContent();
  }

  if (finalRows.length > 0) {
    milestonesSheet.getRange(dataStartRow, 1, finalRows.length, writeWidth).setValues(finalRows);
  }
}

function cleanupRemovedMilestoneTodoistTasks_(removedRows, baseMilestoneColCount) {
  if (!removedRows || removedRows.length === 0) return;
  if (typeof todoistCloseTask_ !== 'function') return;

  removedRows.forEach(function(row) {
    try {
      var meta = extractMilestoneMeta_(row, baseMilestoneColCount);
      var taskId = (meta.todoist_task_id || '').toString().trim();
      if (!taskId) return;

      if (typeof prependSystemCleanupPrefixToTodoistTask_ === 'function') {
        prependSystemCleanupPrefixToTodoistTask_(taskId);
      }

      todoistCloseTask_(taskId);
    } catch (err) {
      Logger.log('[TodoistSync] removed milestone close failed: %s', err && err.message ? err.message : err);
    }
  });
}

function sliceMilestoneBaseRow_(row, baseColCount) {
  var output = [];
  var width = Math.max(0, Number(baseColCount) || 0);

  for (var i = 0; i < width; i++) {
    output.push(normalizeMilestoneBaseCell_(row[i]));
  }

  return output;
}

function normalizeMilestoneBaseCell_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd');
  }

  return (value === null || value === undefined) ? '' : String(value).trim();
}

function isSameMilestoneBaseRow_(a, b) {
  if (!a || !b) return false;
  if (a.length !== b.length) return false;

  for (var i = 0; i < a.length; i++) {
    if (a[i] !== b[i]) return false;
  }

  return true;
}

function normalizeRowsToWidth_(rows, width) {
  return rows.map(function(row) {
    var normalized = row ? row.slice(0, width) : [];
    while (normalized.length < width) {
      normalized.push('');
    }
    return normalized;
  });
}

function makeMilestoneIdentityKey_(row) {
  if (!row) return '';

  var uid = '';
  if (row.length >= INTERIOR_TASK_UID.MILESTONE_UID_COL) {
    uid = normalizeIdentityPart_(row[INTERIOR_TASK_UID.MILESTONE_UID_COL - 1]);
  }
  if (uid) return 'uid||' + uid;

  if (row.length < 4) return '';
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
  meta.process_mark = read(4) || '';
  meta.hasMeta = !!(meta.todoist_task_id || meta.sync_status || meta.last_synced_at || meta.last_error || meta.process_mark);
  return meta;
}

function applyMilestoneMeta_(row, meta, baseColCount) {
  var output = row.slice();
  var normalizedBase = Math.max(0, Number(baseColCount) || 0);

  while (output.length < normalizedBase + 5) {
    output.push('');
  }

  output[normalizedBase] = meta.todoist_task_id || '';
  output[normalizedBase + 1] = meta.sync_status || '';
  output[normalizedBase + 2] = meta.last_synced_at || '';
  output[normalizedBase + 3] = meta.last_error || '';
  output[normalizedBase + 4] = meta.process_mark || '';
  return output;
}
