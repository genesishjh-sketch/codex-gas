/** Logger.gs
 * sync_logs 기록 모듈.
 */

function logSyncEvent_(payload) {
  var headers = ['timestamp', 'action', 'sheet_name', 'block_start_row', 'representative_row', 'project_unique_id', 'clickup_task_id', 'status', 'message'];
  var sh = ensureSheet_(CLICKUP_SETTINGS.LOG_SHEET_NAME, headers);

  var row = [
    new Date(),
    payload.action || '',
    payload.sheetName || '',
    payload.blockStartRow || '',
    payload.representativeRow || '',
    payload.projectUniqueId || '',
    payload.clickupTaskId || '',
    payload.status || '',
    payload.message || ''
  ];
  sh.appendRow(row);
}
