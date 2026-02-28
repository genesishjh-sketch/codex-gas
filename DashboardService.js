/** 90일 집계 생성/갱신 */
function refreshInteriorDashboard90d() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var milestonesSheet = getOrCreateTargetSheet_(ss, INTERIOR_SYNC_CONFIG.TARGET_MILESTONES_ALIASES, INTERIOR_SYNC_CONFIG.TARGET_MILESTONES, INTERIOR_SYNC_CONFIG.TARGET_HEADERS.milestones);
  var dashboard = getOrCreateTargetSheet_(ss, ['대시보드_핵심지표_90일'], '대시보드_핵심지표_90일', ['지표', '값']);

  var lastRow = milestonesSheet.getLastRow();
  var rows = (lastRow >= 2)
    ? milestonesSheet.getRange(2, 1, lastRow - 1, milestonesSheet.getLastColumn()).getValues()
    : [];

  var now = new Date();
  var from = new Date(now.getTime() - (90 * 24 * 60 * 60 * 1000));

  var in90 = rows.filter(function(row) {
    var plan = row[3];
    var planDate = plan instanceof Date ? plan : (plan ? new Date(plan) : null);
    return planDate && !isNaN(planDate.getTime()) && planDate >= from;
  });

  var delayed = in90.filter(function(row) {
    var plan = row[3];
    var done = row[4];
    var planDate = plan instanceof Date ? plan : (plan ? new Date(plan) : null);
    var doneDate = done instanceof Date ? done : (done ? new Date(done) : null);
    return planDate && !doneDate && planDate < now;
  }).length;

  var doneCount = in90.filter(function(row) { return !!row[4]; }).length;
  var total = in90.length;
  var rate = total > 0 ? Math.round((doneCount / total) * 100) : 0;

  var kpi = [
    ['최근90일 작업수', total],
    ['최근90일 완료수', doneCount],
    ['최근90일 완료율(%)', rate],
    ['지연 작업수', delayed]
  ];

  if (dashboard.getLastRow() >= 2) {
    dashboard.getRange(2, 1, dashboard.getLastRow() - 1, 2).clearContent();
  }
  dashboard.getRange(2, 1, kpi.length, 2).setValues(kpi);

  return { total: total, done: doneCount, delayed: delayed, completionRate: rate };
}
