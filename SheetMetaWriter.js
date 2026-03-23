/** SheetMetaWriter.gs
 * 대표행 메타 기록 모듈.
 */

function writeItemSyncKeys_(sheet, representativeRow, projectUniqueId, settings) {
  var homeCol = colToNumber_((settings.basic && settings.basic.HOME_ITEM_KEY_COLUMN) || 'Y');
  var supportCol = colToNumber_((settings.basic && settings.basic.SUPPORT_ITEM_KEY_COLUMN) || 'Z');

  var homeCodes = ['HOME_MEASURE', 'HOME_CONSULT', 'HOME_DESIGN', 'HOME_SHOPPING_LIST', 'HOME_SETTING'];
  var supportCodes = ['SUPPORT_CONSTRUCTION_DONE', 'SUPPORT_FLORAL_DRAFT', 'SUPPORT_FIRST_SETTING', 'SUPPORT_FINAL_SETTING', 'SUPPORT_FLORAL_SCHEDULE'];

  var homeValues = homeCodes.map(function(code) { return [buildItemSyncKey_(projectUniqueId, code)]; });
  var supportValues = supportCodes.map(function(code) { return [buildItemSyncKey_(projectUniqueId, code)]; });

  sheet.getRange(representativeRow, homeCol, 5, 1).setValues(homeValues);
  sheet.getRange(representativeRow, supportCol, 5, 1).setValues(supportValues);
}

function writeClickUpStatus_(sheet, representativeRow, status, settings, extra) {
  var basic = settings.basic || {};
  var taskIdCol = colToNumber_(basic.CLICKUP_TASK_ID_COLUMN || 'U');
  var statusCol = colToNumber_(basic.CLICKUP_CREATE_STATUS_COLUMN || 'AC');

  var urlOffset = Number(basic.CLICKUP_TASK_URL_ROW_OFFSET || 1);
  var syncOffset = Number(basic.SYNC_STATUS_ROW_OFFSET || 2);
  var syncedAtOffset = Number(basic.LAST_SYNCED_AT_ROW_OFFSET || 3);
  var checkOffset = Number(basic.CREATE_CHECK_ROW_OFFSET || 4);

  var now = new Date();

  if (extra && extra.taskId) {
    sheet.getRange(representativeRow, taskIdCol).setValue(extra.taskId);
  }
  if (extra && extra.taskUrl) {
    sheet.getRange(representativeRow + urlOffset, taskIdCol).setValue(extra.taskUrl);
  }

  sheet.getRange(representativeRow + syncOffset, taskIdCol).setValue(status);
  sheet.getRange(representativeRow + syncedAtOffset, taskIdCol).setValue(now);
  sheet.getRange(representativeRow, statusCol).setValue(status);

  if (status === 'CREATED') {
    sheet.getRange(representativeRow + checkOffset, taskIdCol).setValue(true);
  }
}
