/** ClickUpSync.gs
 * V1: 블록 1개 생성 전용 ClickUp 동기화.
 */

function runClickUpCreateForNewBlock_(summary) {
  var ui = SpreadsheetApp.getUi();
  clearClickUpSettingsCache_();
  var settings = getClickUpSettings_();

  var missing = validateClickUpRequiredSettings_(settings);
  if (missing.length > 0) {
    ui.alert(
      '⚠️ ClickUp 설정값이 비어 있습니다.\n' +
      '"clickup settings" 시트에서 아래 항목을 입력 후 다시 실행해주세요.\n- ' + missing.join('\n- ')
    );
    return;
  }

  var sheetName = stringValue_((settings.basic && settings.basic.SOURCE_SHEET_NAME) || CONFIG.SHEET_NAME || '통합관리시트');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('원본 시트를 찾을 수 없습니다: ' + sheetName);

  var startRow = resolveTargetBlockStartRowForClickUp_(summary, settings);
  if (!startRow) {
    ui.alert('ℹ️ ClickUp 생성 대상 블록을 찾지 못했습니다.');
    return;
  }

  var projectData = parseProjectBlock_(sheet, startRow, settings);
  var repRow = projectData.representativeRow;
  var taskIdCol = colToNumber_((settings.basic && settings.basic.CLICKUP_TASK_ID_COLUMN) || 'U');
  var statusCol = colToNumber_((settings.basic && settings.basic.CLICKUP_CREATE_STATUS_COLUMN) || 'AC');

  var existingTaskId = stringValue_(sheet.getRange(repRow, taskIdCol).getDisplayValue());
  var existingCreateStatus = stringValue_(sheet.getRange(repRow, statusCol).getDisplayValue()).toUpperCase();

  // idempotency 정책
  if (existingTaskId || existingCreateStatus === 'CREATED') {
    writeClickUpStatus_(sheet, repRow, 'SKIPPED', settings, {});
    logSyncEvent_({
      action: 'CLICKUP_CREATE',
      sheetName: sheet.getName(),
      blockStartRow: startRow,
      representativeRow: repRow,
      projectUniqueId: stringValue_(sheet.getRange(repRow, colToNumber_((settings.basic && settings.basic.PROJECT_UNIQUE_ID_COLUMN) || 'X')).getDisplayValue()),
      clickupTaskId: existingTaskId,
      status: 'SKIPPED',
      message: '이미 생성된 프로젝트로 판단하여 재생성하지 않음'
    });
    ui.alert('ℹ️ 이미 ClickUp 생성 완료된 프로젝트입니다.\n(재생성하지 않고 SKIPPED 처리)');
    return;
  }

  var projectUniqueId = ensureProjectUniqueId_(sheet, projectData, settings);
  writeItemSyncKeys_(sheet, repRow, projectUniqueId, settings);

  var client = createClickUpClient_(settings);
  var enableDueDate = String(settings.target.CLICKUP_ENABLE_DUE_DATE || 'TRUE').toUpperCase() !== 'FALSE';
  var parentStatus = stringValue_(settings.target.CLICKUP_DEFAULT_PARENT_STATUS);
  var subtaskStatus = stringValue_(settings.target.CLICKUP_DEFAULT_SUBTASK_STATUS);
  var completedStatus = stringValue_(settings.target.CLICKUP_COMPLETED_STATUS);

  var parentTitle = buildProjectTitleFallback_(projectData, projectUniqueId);
  var parentDesc = buildParentTaskDescription_(projectData, projectUniqueId);

  try {
    var parentTask = client.createParentTask(
      parentTitle,
      parentDesc,
      parentStatus,
      enableDueDate ? toClickUpDueDateMs_(projectData.listDeadline, projectData.timeZone) : null
    );

    var subtaskPayloads = buildSubtaskPayloads_(sheet, startRow, projectUniqueId, settings);
    for (var i = 0; i < subtaskPayloads.length; i++) {
      var p = subtaskPayloads[i];
      var desiredStatus = subtaskStatus;
      if (p.doneDateRaw && completedStatus) desiredStatus = completedStatus;

      try {
        client.createSubtask(
          parentTask.id,
          p.name,
          p.description,
          desiredStatus,
          enableDueDate ? toClickUpDueDateMs_(p.dueDateRaw, projectData.timeZone) : null
        );
      } catch (subErr) {
        // 완료 상태 적용 실패 등은 기본 상태로 재시도(전체 실패 방지)
        client.createSubtask(
          parentTask.id,
          p.name,
          p.description,
          subtaskStatus,
          enableDueDate ? toClickUpDueDateMs_(p.dueDateRaw, projectData.timeZone) : null
        );
      }
    }

    writeClickUpStatus_(sheet, repRow, 'CREATED', settings, {
      taskId: parentTask.id,
      taskUrl: parentTask.url || ''
    });

    logSyncEvent_({
      action: 'CLICKUP_CREATE',
      sheetName: sheet.getName(),
      blockStartRow: startRow,
      representativeRow: repRow,
      projectUniqueId: projectUniqueId,
      clickupTaskId: parentTask.id,
      status: 'CREATED',
      message: '부모 Task + Subtask 10개 생성 완료'
    });
  } catch (e) {
    var errMsg = e && e.message ? e.message : String(e);
    writeClickUpStatus_(sheet, repRow, 'ERROR', settings, {});
    logSyncEvent_({
      action: 'CLICKUP_CREATE',
      sheetName: sheet.getName(),
      blockStartRow: startRow,
      representativeRow: repRow,
      projectUniqueId: projectUniqueId,
      clickupTaskId: '',
      status: 'ERROR',
      message: errMsg
    });
    if (errMsg.indexOf('List ID invalid') >= 0) {
      ui.alert(
        '❌ ClickUp 생성 실패\n' +
        'CLICKUP_LIST_ID가 올바르지 않습니다.\n' +
        'clickup settings 시트의 CLICKUP_LIST_ID에 숫자 ID 또는 List URL을 넣어주세요.\n' +
        '(현재값: ' + stringValue_(settings.target.CLICKUP_LIST_ID) + ')'
      );
    } else {
      ui.alert('❌ ClickUp 생성 실패\n' + errMsg);
    }
  }
}
