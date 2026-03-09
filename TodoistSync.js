/** milestones ↔ Todoist 동기화 로직 */
function onMilestonesEditInstallable(e) {
  try {
    if (!e || !e.range) return;
    var settings = getTodoistSyncSettings_();
    if (!settings.realtime_sync) return;

    var sheet = e.range.getSheet();
    if (sheet.getName() !== settings.sync_target_sheet) return;
    if (e.range.getRow() <= 1) return;

    ensureMilestonesSyncColumns_(sheet);
    syncMilestoneRowByRowNumber_(sheet, e.range.getRow(), settings);
  } catch (err) {
    Logger.log('[TodoistSync] onEdit 오류: %s', err && err.message ? err.message : err);
    throw err;
  }
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
  var lastRow = sheet.getLastRow();
  for (var row = 2; row <= lastRow; row++) {
    // L열 처리 마커가 있으면 이미 점검/동기화 완료로 보고 재검사하지 않습니다.
    if (isAlreadyProcessedRow_(sheet, row)) {
      Logger.log('[TodoistSync] row=%s skip: already processed marker in L column', row);
      continue;
    }
    syncMilestoneRowByRowNumber_(sheet, row, settings);
  }
}

function runTodoistMilestonesFullSyncByTrigger() {
  runTodoistMilestonesFullSync();
}

function syncMilestoneRowByRowNumber_(sheet, row, settings) {
  var sectionMap = getSectionMappingMap_();
  var managerMap = getManagerMappingMap_();
  var rowObj = getMilestoneRowObject_(sheet, row);

  Logger.log('[TodoistSync] row=%s data=%s', row, JSON.stringify(rowObj));

  var validate = validateSyncCondition_(rowObj, settings, sectionMap);
  if (!validate.ok) {
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

  var sectionId = getTodoistSectionIdBySection_(rowObj.section, sectionMap);
  if (!sectionId) {
    setSyncResult_(sheet, row, TODOIST_SYNC.STATUS.ERROR, '', 'section 매핑 없음: ' + rowObj.section);
    return;
  }

  var payload = {
    project_id: settings.todoist_project_id,
    section_id: sectionId,
    content: taskContent
  };

  if (dueValue) payload.due_date = formatDateForTodoist_(dueValue);
  appendAssigneeIfEnabled_(payload, settings, rowObj, managerMap);
  appendLabelsIfEnabled_(payload, settings, context);
  appendDescriptionIfEnabled_(payload, settings, context);

  try {
    var result;
    if (!rowObj.todoist_task_id) {
      result = todoistCreateTask_(payload);
      updateTaskId_(sheet, row, result.id);
      setSyncResult_(sheet, row, TODOIST_SYNC.STATUS.CREATED, '', '');
      markRowProcessed_(sheet, row, 'Todoist 동기화완료');
    } else {
      todoistUpdateTask_(rowObj.todoist_task_id, payload);
      setSyncResult_(sheet, row, TODOIST_SYNC.STATUS.UPDATED, '', '');
      markRowProcessed_(sheet, row, 'Todoist 업데이트완료');
    }
  } catch (err) {
    setSyncResult_(sheet, row, TODOIST_SYNC.STATUS.ERROR, '', err.message || String(err));
  }
}

function validateSyncCondition_(rowObj, settings, sectionMap) {
  if (!settings.todoist_project_id) return { ok: false, reason: 'todoist_project_id가 비어 있음' };
  if (!rowObj.plan_date) return { ok: false, reason: 'plan_date 비어 있음' };
  if (settings.exclude_done && rowObj.done_date) return { ok: false, reason: 'done_date가 있어 제외됨' };
  if (!getTodoistSectionIdBySection_(rowObj.section, sectionMap)) return { ok: false, reason: 'section 매핑 없음' };
  return { ok: true };
}

function getMilestoneRowObject_(sheet, row) {
  var values = sheet.getRange(row, 1, 1, 12).getValues()[0];
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
    process_mark: values[11]
  };
}

function appendAssigneeIfEnabled_(payload, settings, rowObj, managerMap) {
  if (!settings.use_assignee) return;
  var mapping = getTodoistAssigneeByManager_(rowObj.manager, managerMap);

  if (!mapping) {
    if (TODOIST_SYNC.ASSIGNEE_POLICY.ERROR_IF_NOT_FOUND) {
      throw new Error('담당자 매핑 없음 또는 inactive: ' + rowObj.manager);
    }
    return;
  }

  payload.assignee_id = mapping.todoist_user_id;
}

function appendDescriptionIfEnabled_(payload, settings) {
  if (!settings.use_description) return;
  payload.description = '';
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
  var requiredCols = Math.max(expected.length, TODOIST_SYNC.PROCESS_MARK.COLUMN_INDEX);
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
}

function setSyncResult_(sheet, row, status, reason, errorText) {
  sheet.getRange(row, 8, 1, 3).setValues([[status + (reason ? ' (' + reason + ')' : ''), new Date(), errorText || '']]);
}

function updateTaskId_(sheet, row, taskId) {
  if (!taskId) return;
  sheet.getRange(row, 7).setValue(String(taskId));
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
