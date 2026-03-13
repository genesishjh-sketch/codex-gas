/** milestones ↔ Todoist 동기화 로직 */
function onMilestonesEditInstallable(e) {
  // DB 동기화 기반 큐 처리로 전환되어 onEdit 실시간 동기화는 비활성화합니다.
  Logger.log('[TodoistSync] onEdit 실시간 동기화 비활성화: runTodoistPendingQueueSync를 사용하세요.');
}

function syncSelectedMilestoneRowToTodoist() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var settings = getTodoistSyncSettings_();
  if (sheet.getName() !== settings.sync_target_sheet) {
    throw new Error('현재 활성 시트가 sync_target_sheet와 다릅니다.');
  }

  var row = sheet.getActiveRange().getRow();
  if (row <= 1) throw new Error('헤더 행은 동기화 대상이 아닙니다.');
  ensureMilestonesSyncColumns_(sheet);
  syncMilestoneRowByRowNumber_(sheet, row, settings);
}

function runTodoistMilestonesFullSync() {
  var settings = getTodoistSyncSettings_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(settings.sync_target_sheet);
  if (!sheet) throw new Error('동기화 대상 시트를 찾을 수 없습니다: ' + settings.sync_target_sheet);

  ensureMilestonesSyncColumns_(sheet);

  var fullSyncSettings = TODOIST_SYNC.FULL_SYNC || {};
  var maxRuntimeMs = Number(fullSyncSettings.MAX_RUNTIME_MS) || 330000;
  var cursorPropertyKey = fullSyncSettings.CURSOR_PROPERTY_KEY || 'TODOIST_FULL_SYNC_LAST_ROW';
  var scriptProps = PropertiesService.getScriptProperties();

  var lastRow = sheet.getLastRow();
  var defaultStartRow = 2;
  var savedCursor = parseInt(scriptProps.getProperty(cursorPropertyKey) || '', 10);
  var startRow = (savedCursor >= defaultStartRow && savedCursor <= lastRow) ? savedCursor : defaultStartRow;

  var startedAt = Date.now();
  var completed = true;

  for (var row = startRow; row <= lastRow; row++) {
    if ((Date.now() - startedAt) >= maxRuntimeMs) {
      scriptProps.setProperty(cursorPropertyKey, String(row));
      Logger.log('[TodoistSync] runtime guard reached at row=%s, cursor saved', row);
      completed = false;
      break;
    }

    // L열 처리 마커가 있으면 이미 점검/동기화 완료로 보고 재검사하지 않습니다.
    if (isAlreadyProcessedRow_(sheet, row)) {
      Logger.log('[TodoistSync] row=%s skip: already processed marker in L column', row);
      continue;
    }
    syncMilestoneRowByRowNumber_(sheet, row, settings);
  }

  if (completed) {
    scriptProps.deleteProperty(cursorPropertyKey);
  }
}

function runTodoistMilestonesFullSyncByTrigger() {
  runTodoistMilestonesFullSync();
}


function runTodoistPendingQueueSyncByTrigger() {
  runTodoistPendingQueueSync();
}

function runTodoistCompletionMirrorHourlyByTrigger() {
  runTodoistCompletionMirrorToSource_();
}

function runTodoistCompletionMirrorOnOpen_() {
  var cfg = (TODOIST_SYNC && TODOIST_SYNC.COMPLETION_MIRROR) || {};
  var minIntervalSec = Number(cfg.ON_OPEN_MIN_INTERVAL_SEC) || 120;
  var props = PropertiesService.getDocumentProperties();
  var key = 'TODOIST_COMPLETION_MIRROR_LAST_ONOPEN_AT';
  var now = Date.now();
  var last = Number(props.getProperty(key) || 0);
  if (last && (now - last) < (minIntervalSec * 1000)) return;

  runTodoistCompletionMirrorToSource_();
  props.setProperty(key, String(now));
}

function runTodoistPendingQueueSync() {
  var settings = getTodoistSyncSettings_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(settings.sync_target_sheet);
  if (!sheet) throw new Error('동기화 대상 시트를 찾을 수 없습니다: ' + settings.sync_target_sheet);

  ensureMilestonesSyncColumns_(sheet);

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var markerCol = TODOIST_SYNC.PROCESS_MARK.COLUMN_INDEX;
  var markerValues = sheet.getRange(2, markerCol, lastRow - 1, 1).getDisplayValues();

  for (var idx = 0; idx < markerValues.length; idx++) {
    var marker = (markerValues[idx][0] || '').toString().trim();
    if (marker !== '신규' && marker !== '수정') continue;

    var row = idx + 2;
    syncMilestoneRowByRowNumber_(sheet, row, settings);

    var status = sheet.getRange(row, 8).getDisplayValue();
    if (status && status.indexOf(TODOIST_SYNC.STATUS.ERROR) === 0) continue;

    sheet.getRange(row, markerCol).setValue('반영');
  }

  runTodoistCompletionMirrorToSource_();
}

function scheduleTodoistPendingQueueSyncFallback_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === TODOIST_SYNC.PENDING_QUEUE_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  var burstCfg = (TODOIST_SYNC && TODOIST_SYNC.COMPLETION_MIRROR) || {};
  var burstWindow = Math.max(2, Number(burstCfg.BURST_WINDOW_MINUTES) || 20);
  var burstStep = Math.max(1, Number(burstCfg.BURST_STEP_MINUTES) || 2);

  for (var minute = burstStep; minute <= burstWindow; minute += burstStep) {
    ScriptApp.newTrigger(TODOIST_SYNC.PENDING_QUEUE_TRIGGER_HANDLER)
      .timeBased()
      .after(minute * 60 * 1000)
      .create();
  }
}


function syncMilestoneRowByRowNumber_(sheet, row, settings) {
  var sectionMap = getSectionMappingMap_();
  var managerMap = getManagerMappingMap_();
  var stepProjectRules = getStepProjectMappingRules_();
  var rowObj = getMilestoneRowObject_(sheet, row);
  var effectiveManager = resolveManagerByStepName_(rowObj.step_name, rowObj.manager);
  var resolvedProjectId = resolveTodoistProjectIdByStepName_(rowObj.step_name, settings.todoist_project_id, stepProjectRules);

  Logger.log('[TodoistSync] row=%s data=%s', row, JSON.stringify(rowObj));

  var validate = validateSyncCondition_(rowObj, settings, resolvedProjectId);
  if (!validate.ok) {
    if (validate.shouldCloseTodoistTask && rowObj.todoist_task_id) {
      try {
        prependSystemCleanupPrefixToTodoistTask_(rowObj.todoist_task_id);
        todoistCloseTask_(rowObj.todoist_task_id);
        setSyncResult_(sheet, row, TODOIST_SYNC.STATUS.UPDATED, 'Todoist 작업 완료 처리: ' + validate.reason, '');
        setSyncSource_(sheet, row, 'sheet');
        markRowProcessed_(sheet, row, 'Todoist 완료처리');
      } catch (closeErr) {
        setSyncResult_(sheet, row, TODOIST_SYNC.STATUS.ERROR, '', closeErr.message || String(closeErr));
      }
      return;
    }

    setSyncResult_(sheet, row, TODOIST_SYNC.STATUS.SKIPPED, validate.reason, '');
    return;
  }

  var context = buildTemplateContextFromMilestoneRow_(rowObj);
  var taskContent = renderSafeTemplate_(settings.task_title_template, context);
  var dueValue = context[settings.due_date_field] || '';

  if (!taskContent) {
    setSyncResult_(sheet, row, TODOIST_SYNC.STATUS.ERROR, '', 'task_title_template 결과가 비어 있습니다.');
    return;
  }

  var sectionId = getTodoistSectionIdBySection_(rowObj.section, sectionMap, resolvedProjectId);

  var payload = {
    project_id: resolvedProjectId,
    content: taskContent
  };
  if (sectionId) payload.section_id = sectionId;

  if (dueValue) payload.due_date = formatDateForTodoist_(dueValue);
  appendTaskUidMetaToPayload_(payload, rowObj.task_uid);
  appendAssigneeIfEnabled_(payload, settings, effectiveManager, managerMap, resolvedProjectId);
  appendLabelsIfEnabled_(payload, settings, context);
  appendDescriptionIfEnabled_(payload, settings, context);

  try {
    var result;
    if (!rowObj.todoist_task_id) {
      var existingTask = findExistingTodoistTaskByUid_(resolvedProjectId, rowObj.task_uid);
      if (existingTask && existingTask.id) {
        updateTaskId_(sheet, row, existingTask.id);
        rowObj.todoist_task_id = existingTask.id;
      }
    }

    if (!rowObj.todoist_task_id) {
      result = todoistCreateTask_(payload);
      updateTaskId_(sheet, row, result.id);
      setSyncResult_(sheet, row, TODOIST_SYNC.STATUS.CREATED, '', '');
      setSyncSource_(sheet, row, 'sheet');
      markRowProcessed_(sheet, row, 'Todoist 동기화완료');
    } else {
      if (!rowObj.done_date) {
        try { todoistReopenTask_(rowObj.todoist_task_id); } catch (reopenErr) {}
      }
      todoistUpdateTask_(rowObj.todoist_task_id, payload);
      setSyncResult_(sheet, row, TODOIST_SYNC.STATUS.UPDATED, '', '');
      setSyncSource_(sheet, row, 'sheet');
      markRowProcessed_(sheet, row, 'Todoist 업데이트완료');
    }
  } catch (err) {
    setSyncResult_(sheet, row, TODOIST_SYNC.STATUS.ERROR, '', err.message || String(err));
  }
}

function findExistingTodoistTaskByUid_(projectId, taskUid) {
  var normalizedUid = (taskUid || '').toString().trim();
  if (!normalizedUid) return null;

  try {
    return todoistFindActiveTaskByTaskUid_(projectId, normalizedUid);
  } catch (e) {
    Logger.log('[TodoistSync] find by uid failed project=%s uid=%s err=%s', projectId, normalizedUid, e && e.message ? e.message : e);
    return null;
  }
}


function prependSystemCleanupPrefixToTodoistTask_(taskId) {
  var prefix = '[시스템정리] ';
  var task = todoistGetTask_(taskId) || {};
  var currentContent = (task.content || '').toString().trim();
  if (!currentContent) return;
  if (currentContent.indexOf(prefix) === 0) return;
  todoistUpdateTask_(taskId, { content: prefix + currentContent });
}

function validateSyncCondition_(rowObj, settings, resolvedProjectId) {
  if (!resolvedProjectId) return { ok: false, reason: 'todoist_project_id가 비어 있음(step_name 매핑 포함)' };
  if (!rowObj.plan_date) {
    return { ok: false, reason: 'plan_date 비어 있음', shouldCloseTodoistTask: true };
  }
  if (settings.exclude_done && rowObj.done_date) {
    return { ok: false, reason: 'done_date가 있어 제외됨', shouldCloseTodoistTask: true };
  }
  return { ok: true };
}

function getMilestoneRowObject_(sheet, row) {
  var values = sheet.getRange(row, 1, 1, 13).getValues()[0];
  return {
    project_code: values[0],
    section: values[1],
    step_name: values[2],
    plan_date: values[3],
    done_date: values[4],
    manager: values[5],
    todoist_task_id: values[6],
    sync_status: values[7],
    last_synced_at: values[8],
    last_error: values[9],
    process_mark: values[10],
    sync_source: values[11],
    task_uid: values[12]
  };
}

function appendAssigneeIfEnabled_(payload, settings, managerName, managerMap, projectIdForAssignee) {
  if (!settings.use_assignee) return;
  var mapping = getTodoistAssigneeByManager_(managerName, managerMap);

  if (!mapping) {
    if (TODOIST_SYNC.ASSIGNEE_POLICY.ERROR_IF_NOT_FOUND) {
      throw new Error('담당자 매핑 없음 또는 inactive: ' + managerName);
    }
    return;
  }

  var assigneeProjectId = (projectIdForAssignee || settings.todoist_project_id || '').toString().trim();
  var assigneeId = resolveAssigneeIdFromMapping_(mapping, assigneeProjectId);
  if (!assigneeId) {
    if (TODOIST_SYNC.ASSIGNEE_POLICY.ERROR_IF_NOT_FOUND) {
      throw new Error('담당자 ID 해석 실패(manager=' + managerName + ', email=' + (mapping.todoist_user_email || '') + ')');
    }
    return;
  }

  payload.assignee_id = assigneeId;
}

function resolveManagerByStepName_(stepName, originalManagerName) {
  var step = (stepName || '').toString().trim();
  if (!step) return originalManagerName;

  var policy = TODOIST_SYNC.FORCED_MANAGER_MAPPING || {};
  var forceManager = (policy.MANAGER_NAME || '').toString().trim();
  var stepNames = policy.STEP_NAMES || [];
  if (!forceManager || !stepNames || !stepNames.length) return originalManagerName;

  var shouldForce = stepNames.some(function(candidate) {
    return (candidate || '').toString().trim() === step;
  });

  return shouldForce ? forceManager : originalManagerName;
}

function resolveTodoistProjectIdByStepName_(stepName, defaultProjectId, rules) {
  var step = (stepName || '').toString().trim();
  var matchedProjectId = '';

  (rules || []).some(function(rule) {
    if (!matchStepNameRule_(step, rule)) return false;
    matchedProjectId = (rule.todoist_project_id || '').toString().trim();
    return !!matchedProjectId;
  });

  if (matchedProjectId) return matchedProjectId;
  return (defaultProjectId || '').toString().trim();
}

function matchStepNameRule_(stepName, rule) {
  if (!rule) return false;

  var step = (stepName || '').toString();
  var pattern = (rule.pattern || '').toString();
  if (!pattern) return false;

  if (rule.match_type === 'exact') {
    return step.trim() === pattern.trim();
  }

  if (rule.match_type === 'contains') {
    return step.indexOf(pattern) >= 0;
  }

  if (rule.match_type === 'regex') {
    try {
      var re = new RegExp(pattern);
      return re.test(step);
    } catch (err) {
      Logger.log('[TodoistSync] invalid step-project regex pattern: %s', pattern);
      return false;
    }
  }

  return false;
}

function resolveAssigneeIdFromMapping_(mapping, projectId) {
  var explicitId = (mapping.todoist_user_id || '').toString().trim();
  if (explicitId) return explicitId;

  var email = (mapping.todoist_user_email || '').toString().trim();
  if (!email) return '';

  return todoistFindCollaboratorIdByEmail_(projectId, email);
}

function appendDescriptionIfEnabled_(payload, settings) {
  if (!settings.use_description) return;
  payload.description = payload.description || '';
}

function appendTaskUidMetaToPayload_(payload, taskUid) {
  var uid = (taskUid || '').toString().trim();
  if (!uid) return;
  var marker = 'meta_task_uid:' + uid;
  var desc = (payload.description || '').toString();
  if (desc.indexOf(marker) >= 0) return;
  payload.description = desc ? (desc + '\n' + marker) : marker;
}

function appendLabelsIfEnabled_(payload, settings, context) {
  if (!settings.use_labels) return;
  if (!settings.label_template) return;

  var rendered = renderSafeTemplate_(settings.label_template, context);
  var labels = rendered.split(',').map(function(v) { return v.trim(); }).filter(function(v) { return !!v; });
  if (labels.length > 0) payload.labels = labels;
}

function ensureMilestonesSyncColumns_(sheet) {
  var expected = TODOIST_SYNC.MILESTONE_HEADERS.concat(TODOIST_SYNC.SYNC_HEADERS);
  var requiredCols = Math.max(expected.length, TODOIST_SYNC.PROCESS_MARK.COLUMN_INDEX, INTERIOR_TASK_UID.MILESTONE_UID_COL);
  if (sheet.getMaxColumns() < requiredCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredCols - sheet.getMaxColumns());
  }

  var current = sheet.getRange(1, 1, 1, expected.length).getDisplayValues()[0];
  var needsWrite = false;
  for (var i = 0; i < expected.length; i++) {
    if ((current[i] || '').toString().trim() !== expected[i]) {
      needsWrite = true;
      break;
    }
  }
  if (needsWrite) sheet.getRange(1, 1, 1, expected.length).setValues([expected]);

  sheet.getRange(1, TODOIST_SYNC.PROCESS_MARK.COLUMN_INDEX).setValue('process_mark');
  sheet.getRange(1, TODOIST_SYNC.SYNC_SOURCE_COLUMN_INDEX).setValue('sync_source');
  var uidHeaderCell = sheet.getRange(1, INTERIOR_TASK_UID.MILESTONE_UID_COL);
  uidHeaderCell.setValue(INTERIOR_TASK_UID.MILESTONE_UID_HEADER_TEXT);
  uidHeaderCell.setNote(INTERIOR_TASK_UID.MILESTONE_UID_WARNING_TEXT);
}

function setSyncResult_(sheet, row, status, reason, errorText) {
  sheet.getRange(row, 8, 1, 3).setValues([[status + (reason ? ' (' + reason + ')' : ''), new Date(), errorText || '']]);
}

function updateTaskId_(sheet, row, taskId) {
  sheet.getRange(row, 7).setValue(taskId ? String(taskId) : '');
}

function setSyncSource_(sheet, row, source) {
  sheet.getRange(row, TODOIST_SYNC.SYNC_SOURCE_COLUMN_INDEX).setValue((source || '').toString().trim());
}


function isAlreadyProcessedRow_(sheet, row) {
  var marker = sheet.getRange(row, TODOIST_SYNC.PROCESS_MARK.COLUMN_INDEX).getDisplayValue();
  if (!marker) return false;

  var normalized = marker.toString().trim().toLowerCase();
  if (!normalized) return false;

  return TODOIST_SYNC.PROCESS_MARK.SKIP_KEYWORDS.some(function(keyword) {
    return normalized.indexOf(keyword.toLowerCase()) >= 0;
  });
}

function markRowProcessed_(sheet, row, message) {
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheet.getRange(row, TODOIST_SYNC.PROCESS_MARK.COLUMN_INDEX).setValue((message || '동기화완료') + ' | ' + timestamp);
}

function formatDateForTodoist_(value) {
  if (!value) return '';
  var dateObj = (value instanceof Date) ? value : new Date(value);
  if (isNaN(dateObj.getTime())) return '';
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function runTodoistCompletionMirrorToSource_() {
  var settings = getTodoistSyncSettings_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var milestonesSheet = ss.getSheetByName(settings.sync_target_sheet);
  var sourceSheet = getSheetByAliases_(ss, INTERIOR_SYNC_CONFIG.SOURCE_SHEET_ALIASES);
  if (!milestonesSheet || !sourceSheet) return;

  var uidMap = buildSourceUidRowMap_(sourceSheet);
  var lastRow = milestonesSheet.getLastRow();
  if (lastRow < 2) return;

  var rows = milestonesSheet.getRange(2, 1, lastRow - 1, 13).getValues();
  rows.forEach(function(values, idx) {
    var rowNum = idx + 2;
    var taskId = (values[6] || '').toString().trim();
    var doneDate = values[4];
    var uid = (values[12] || '').toString().trim();
    if (!taskId || !uid || doneDate) return;

    var task;
    var isCompleted = false;
    try {
      task = todoistGetTask_(taskId);
      isCompleted = !!task.is_completed;
    } catch (e) {
      // 완료된 태스크는 /tasks/{id}에서 조회 실패(404)할 수 있어 completed 목록에서 보강 조회
      var completedTask = null;
      try {
        completedTask = todoistGetCompletedTaskByTaskId_(taskId);
      } catch (completedErr) {
        completedTask = null;
      }
      isCompleted = !!completedTask;
    }

    if (!isCompleted) return;

    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    milestonesSheet.getRange(rowNum, 5).setValue(today);
    setSyncSource_(milestonesSheet, rowNum, 'todoist');

    var hit = uidMap[uid];
    if (hit) sourceSheet.getRange(hit.row, hit.doneCol).setValue(today);
  });
}

function buildSourceUidRowMap_(sourceSheet) {
  var anchors = collectAnchorRows_(sourceSheet);
  var blockHeight = Math.max(1, (typeof getBlockHeight_ === 'function') ? getBlockHeight_(sourceSheet) : 9);
  var out = {};

  anchors.forEach(function(anchorRow) {
    var start = Math.max(1, anchorRow - 7);
    var end = start + blockHeight - 1;

    for (var r = start; r <= end; r++) {
      var hasHome = !!(sourceSheet.getRange(r, 7).getDisplayValue() || sourceSheet.getRange(r, 8).getDisplayValue() || sourceSheet.getRange(r, 9).getDisplayValue());
      var hasConstruction = !!(sourceSheet.getRange(r, 13).getDisplayValue() || sourceSheet.getRange(r, 14).getDisplayValue() || sourceSheet.getRange(r, 15).getDisplayValue());

      if (hasHome) {
        var homeUid = (sourceSheet.getRange(r, INTERIOR_TASK_UID.SOURCE_HOME_UID_COL).getDisplayValue() || '').toString().trim();
        if (homeUid) out[homeUid] = { row: r, doneCol: 9 };
      }
      if (hasConstruction) {
        var constructionUid = (sourceSheet.getRange(r, INTERIOR_TASK_UID.SOURCE_CONSTRUCTION_UID_COL).getDisplayValue() || '').toString().trim();
        if (constructionUid) out[constructionUid] = { row: r, doneCol: 15 };
      }
    }
  });

  return out;
}
