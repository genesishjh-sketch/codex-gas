/** ProjectId.gs
 * project_unique_id 생성/보존.
 */

function ensureProjectUniqueId_(sheet, projectData, settings) {
  var repRow = projectData.representativeRow;
  var projectIdCol = colToNumber_((settings.basic && settings.basic.PROJECT_UNIQUE_ID_COLUMN) || 'X');
  var existing = stringValue_(sheet.getRange(repRow, projectIdCol).getDisplayValue());
  if (existing) return existing;

  var titleLoc = parseFirstParenText_(projectData.projectTitle);
  var location = titleLoc || parseFirstParenText_(projectData.address) || '현장';
  var customer = projectData.customerName || '고객';

  var hashSource = [
    projectData.blockStartRow,
    customer,
    projectData.address,
    formatDateYmd_(projectData.contractDate, projectData.timeZone)
  ].join('|');

  var projectId = 'PJT-' + safeSlug_(location) + '-' + safeSlug_(customer) + '-' + shortHash6_(hashSource);
  sheet.getRange(repRow, projectIdCol).setValue(projectId);
  return projectId;
}

function buildItemSyncKey_(projectUniqueId, itemCode) {
  return projectUniqueId + '__' + itemCode;
}
