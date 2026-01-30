/** CalendarManager.gs */

function generateWeeklyCalendar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var source = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!source) { SpreadsheetApp.getUi().alert("⚠️ 통합관리시트 없음"); return; }

  var cal = ss.getSheetByName(CONFIG.CALENDAR_SHEET_NAME);
  if (!cal) cal = ss.insertSheet(CONFIG.CALENDAR_SHEET_NAME);
  cal.clear();

  var blockHeight = getBlockHeight_(source);
  var lastRow = source.getLastRow();
  if (lastRow < CONFIG.START_ROW) return;

  var now = new Date(); now.setHours(0,0,0,0);
  var day = now.getDay();
  var weekStart = new Date(now); weekStart.setDate(now.getDate() - day);
  var weekEnd = new Date(weekStart); weekEnd.setDate(weekStart.getDate() + 6);

  var tz = Session.getScriptTimeZone();
  var title = "📅 금주 일정표 (일요일 시작)  " +
    Utilities.formatDate(weekStart, tz, "MM/dd") + " ~ " +
    Utilities.formatDate(weekEnd, tz, "MM/dd");

  cal.getRange(1,1).setValue(title).setFontWeight("bold");
  cal.getRange(2,1).setValue("프로젝트").setFontWeight("bold");

  var dateHeaders = [];
  for (var i = 0; i < 7; i++) {
    var d = new Date(weekStart); d.setDate(weekStart.getDate() + i);
    dateHeaders.push(Utilities.formatDate(d, tz, "MM/dd") + " (" + ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"][i] + ")");
  }
  cal.getRange(2,2,1,7).setValues([dateHeaders]).setFontWeight("bold");

  var rows = [];
  for (var r = CONFIG.START_ROW; r <= lastRow; r += blockHeight) {
    var pname = source.getRange(r, CONFIG.POS_NAME.col).getDisplayValue();
    if (!isValidName(pname)) continue;

    var weekMap = {};
    CONFIG.TASK_COLS.forEach(function(tc){
      var labels = source.getRange(r, tc.labelCol, blockHeight, 1).getDisplayValues();
      var dates = source.getRange(r, tc.dateCol, blockHeight, 1).getValues();

      for (var k = 0; k < blockHeight; k++) {
        var label = (labels[k][0] || "").toString().trim();
        var d = dates[k][0];
        if (!label || !(d instanceof Date)) continue;

        var dd = new Date(d); dd.setHours(0,0,0,0);
        if (dd < weekStart || dd > weekEnd) continue;

        var idx = (dd.getTime() - weekStart.getTime()) / (24*3600*1000);
        idx = Math.round(idx);

        var prefix = tc.prefix || "";
        var txt = label;
        if (prefix && txt.indexOf(prefix) !== 0) txt = prefix + txt;

        weekMap[idx] = weekMap[idx] || [];
        weekMap[idx].push(txt);
      }
    });

    if (Object.keys(weekMap).length === 0) continue;

    var line = new Array(8).fill("");
    line[0] = pname;
    for (var di = 0; di < 7; di++) {
      if (weekMap[di] && weekMap[di].length > 0) {
        line[1 + di] = weekMap[di].join("\n");
      }
    }
    rows.push(line);
  }

  if (rows.length === 0) {
    cal.getRange(3,1).setValue("금주 일정 없음");
    return;
  }

  cal.getRange(3,1,rows.length,8).setValues(rows);
  cal.setFrozenRows(2);
  cal.setColumnWidth(1, 320);
  for (var c = 2; c <= 8; c++) cal.setColumnWidth(c, 140);
  cal.getRange(3,1,rows.length,8).setWrap(true);
}
