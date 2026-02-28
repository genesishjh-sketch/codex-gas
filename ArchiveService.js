/** 완료 + N일(기본 30일) 지난 마일스톤을 보관 시트로 이동 */
function archiveCompletedMilestones() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var milestonesSheet = getOrCreateTargetSheet_(ss, INTERIOR_SYNC_CONFIG.TARGET_MILESTONES_ALIASES, INTERIOR_SYNC_CONFIG.TARGET_MILESTONES, INTERIOR_SYNC_CONFIG.TARGET_HEADERS.milestones);
  var archiveName = 'milestones_archive';
  var archiveHeaders = INTERIOR_SYNC_CONFIG.TARGET_HEADERS.milestones.concat(['archived_at']);
  var archiveSheet = getOrCreateTargetSheet_(ss, [archiveName], archiveName, archiveHeaders);

  var lastRow = milestonesSheet.getLastRow();
  if (lastRow < 2) return { archived: 0 };

  var rows = milestonesSheet.getRange(2, 1, lastRow - 1, milestonesSheet.getLastColumn()).getValues();
  var keep = [];
  var archived = [];
  var thresholdDays = getInteriorArchiveAfterDays_();
  var now = new Date();
  var thresholdMs = thresholdDays * 24 * 60 * 60 * 1000;

  rows.forEach(function(row) {
    var done = row[4];
    var doneDate = done instanceof Date ? done : (done ? new Date(done) : null);
    if (doneDate && !isNaN(doneDate.getTime()) && (now.getTime() - doneDate.getTime()) >= thresholdMs) {
      archived.push(row.concat([new Date()]));
    } else {
      keep.push(row);
    }
  });

  if (lastRow >= 2) {
    milestonesSheet.getRange(2, 1, lastRow - 1, milestonesSheet.getLastColumn()).clearContent();
  }
  if (keep.length > 0) {
    milestonesSheet.getRange(2, 1, keep.length, keep[0].length).setValues(keep);
  }

  if (archived.length > 0) {
    var start = archiveSheet.getLastRow() + 1;
    archiveSheet.getRange(start, 1, archived.length, archived[0].length).setValues(archived);
  }

  return { archived: archived.length, thresholdDays: thresholdDays };
}
