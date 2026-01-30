/** DBManager.gs
 * - DB 시트명 자동 생성 + 고정 재사용
 * - 실측/상담/디자인/엑셀/세팅: 예정/완료 저장
 * - ✅ 메뉴 함수: syncDBActive / syncDBFullScan 제공(= 에러 해결)
 */

function sanitizeSheetName_(name) {
  var s = (name || "").toString();
  s = s.replace(/[\\\/\?\*\[\]:]/g, "_");
  s = s.replace(/\s+/g, " ").trim();
  if (s.length > 90) s = s.slice(0, 90).trim();
  if (!s) s = "DB";
  return s;
}

function getOrCreateDbSheetName_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) Config에 지정된 이름이 있으면 우선
  var fixed = (CONFIG.DB_SHEET_NAME || "").toString().trim();
  if (fixed) return fixed;

  var props = PropertiesService.getScriptProperties();
  var key = "PROJECT_DB_SHEET_NAME";
  var saved = props.getProperty(key);

  if (saved && ss.getSheetByName(saved)) return saved;

  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, "yyyyMMdd");

  var base = "DB_" + (CONFIG.SHEET_NAME || "통합관리시트") + "_" + today;
  base = sanitizeSheetName_(base);

  var name = base;
  var n = 2;
  while (ss.getSheetByName(name)) {
    name = sanitizeSheetName_(base + " (" + n + ")");
    n++;
  }

  props.setProperty(key, name);
  return name;
}

function ensureDbSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = getOrCreateDbSheetName_();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 22).setValues([[
      "KEY", "NO", "프로젝트명", "상태",
      "주소", "지도URL",
      "메인폴더", "Before폴더", "시공폴더", "After폴더", "물품리스트",
      "실측_예정", "실측_완료",
      "상담_예정", "상담_완료",
      "디자인_예정", "디자인_완료",
      "엑셀작업_예정", "엑셀작업_완료",
      "세팅_예정", "세팅_완료",
      "UPDATED_AT"
    ]]).setFontWeight("bold");
  }
  return sh;
}

function fmtDate_(v) {
  if (!v) return "";
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "MM/dd");
  return v.toString().trim();
}

/**
 * 블록에서 단계별 예정/완료 뽑기
 * - 라벨: G열(7)
 * - 예정: H열(8)
 * - 완료: I열(9)
 */
function collectStageDates_(sheet, blockStartRow, blockHeight) {
  var labelCol = 7; // G
  var planCol  = 8; // H
  var doneCol  = 9; // I

  var labels = sheet.getRange(blockStartRow, labelCol, blockHeight, 1).getDisplayValues();
  var plans  = sheet.getRange(blockStartRow, planCol,  blockHeight, 1).getValues();
  var dones  = sheet.getRange(blockStartRow, doneCol,  blockHeight, 1).getValues();

  var stages = [
    { key: "MEASURE", keywords: ["실측"] },
    { key: "CONSULT", keywords: ["상담", "줌", "미팅", "zoom"] },
    { key: "DESIGN",  keywords: ["디자인"] },
    { key: "EXCEL",   keywords: ["엑셀", "excel"] },
    { key: "SETTING", keywords: ["세팅", "세팅일자"] }
  ];

  var out = {
    MEASURE_PLAN:"", MEASURE_DONE:"",
    CONSULT_PLAN:"", CONSULT_DONE:"",
    DESIGN_PLAN:"",  DESIGN_DONE:"",
    EXCEL_PLAN:"",   EXCEL_DONE:"",
    SETTING_PLAN:"", SETTING_DONE:""
  };

  function match_(label, keywords) {
    if (!label) return false;
    var t = label.toString().replace(/\s+/g, "").toLowerCase();
    return keywords.some(function(k){
      return t.indexOf(k.toString().replace(/\s+/g,"").toLowerCase()) !== -1;
    });
  }

  for (var i = 0; i < blockHeight; i++) {
    var lab = (labels[i][0] || "").toString().trim();
    if (!lab) continue;

    for (var s = 0; s < stages.length; s++) {
      var st = stages[s];
      var planKey = st.key + "_PLAN";
      var doneKey = st.key + "_DONE";
      if (out[planKey] || out[doneKey]) continue;

      if (match_(lab, st.keywords)) {
        out[planKey] = fmtDate_(plans[i][0]);
        out[doneKey] = fmtDate_(dones[i][0]);
      }
    }
  }

  return out;
}

function makeDbKey_(noVal, nameVal) {
  var no = (noVal || "").toString().trim();
  var nm = (nameVal || "").toString().trim();
  return no + "|" + nm;
}

function loadDbKeyMap_(dbSheet) {
  var map = {};
  var last = dbSheet.getLastRow();
  if (last < 2) return map;

  var keys = dbSheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < keys.length; i++) {
    var k = (keys[i][0] || "").toString().trim();
    if (k) map[k] = i + 2;
  }
  return map;
}

function shouldSkipDbActive_(status) {
  var s = (status || "").toString().trim();
  return (
    s === "완료" ||
    s === "취소" ||
    s === "대기" ||
    s === "세팅 대기" ||
    s === "세팅완료(에비대기)"
  );
}

/**
 * ✅ DB 동기화
 * @param {boolean} includeAll  - true면 전체(완료/취소 포함), false면 진행만
 * @param {boolean} isSilent
 */
function syncProjectDb(includeAll, isSilent) {
  var sheet = getMainSheet_();
  var blockHeight = getBlockHeight_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.START_ROW) return;

  var stopCtl = makeStopController_();
  var db = ensureDbSheet_();
  var keyMap = loadDbKeyMap_(db);

  var appends = [];
  var updates = [];

  for (var r = CONFIG.START_ROW; r <= lastRow; r += blockHeight) {
    if (stopCtl.check(sheet, r)) break;

    var nameVal = sheet.getRange(r, CONFIG.POS_NAME.col).getDisplayValue();
    if (!isValidName(nameVal)) continue;

    var status = getStatus_(sheet, r);
    if (!includeAll && shouldSkipDbActive_(status)) continue;

    var noVal = sheet.getRange(r, CONFIG.POS_NO.col).getDisplayValue();

    // 주소(F4 + 공백 + F6)
    var f4 = sheet.getRange(r + CONFIG.POS_ADDR.row, CONFIG.POS_ADDR.col).getDisplayValue();
    var f6 = sheet.getRange(r + CONFIG.POS_ADDR_EXTRA.row, CONFIG.POS_ADDR_EXTRA.col).getDisplayValue();
    var addrLine = ((f4 || "").toString().trim() + " " + (f6 || "").toString().trim()).replace(/\s+/g, " ").trim();

    var mapUrl = sheet.getRange(r + CONFIG.POS_MAP.row, CONFIG.POS_MAP.col).getDisplayValue();

    // 폴더/파일 링크 (S열 기준, R열은 폴더명)
    var folderUrlCol = (CONFIG && (CONFIG.POS_FOLDER_URL_COL || CONFIG.DRIVE_MARK_COL)) || 19; // S
    var mainCell = sheet.getRange(r + 1, folderUrlCol);
    var sub1Cell = sheet.getRange(r + 2, folderUrlCol);
    var sub2Cell = sheet.getRange(r + 3, folderUrlCol);
    var sub3Cell = sheet.getRange(r + 4, folderUrlCol);
    var mainFolder = mainCell.getDisplayValue();
    var sub1 = sub1Cell.getDisplayValue();
    var sub2 = sub2Cell.getDisplayValue();
    var sub3 = sub3Cell.getDisplayValue();
    // 하이퍼링크만 있는 경우까지 보완
    if (!mainFolder) mainFolder = getUrlFromCell_(mainCell);
    if (!sub1) sub1 = getUrlFromCell_(sub1Cell);
    if (!sub2) sub2 = getUrlFromCell_(sub2Cell);
    if (!sub3) sub3 = getUrlFromCell_(sub3Cell);

    var fileCell = sheet.getRange(r + CONFIG.POS_FILE.row, CONFIG.POS_FILE.col);
    var fileUrl = fileCell.getDisplayValue();
    if (!fileUrl) fileUrl = getUrlFromCell_(fileCell);

    var stages = collectStageDates_(sheet, r, blockHeight);

    var key = makeDbKey_(noVal, nameVal);
    var row = [
      key, noVal, nameVal, status,
      addrLine, mapUrl,
      mainFolder, sub1, sub2, sub3, fileUrl,
      stages.MEASURE_PLAN, stages.MEASURE_DONE,
      stages.CONSULT_PLAN, stages.CONSULT_DONE,
      stages.DESIGN_PLAN, stages.DESIGN_DONE,
      stages.EXCEL_PLAN, stages.EXCEL_DONE,
      stages.SETTING_PLAN, stages.SETTING_DONE,
      new Date()
    ];

    if (keyMap[key]) updates.push({ row: keyMap[key], values: row });
    else appends.push(row);
  }

  for (var i = 0; i < updates.length; i++) {
    db.getRange(updates[i].row, 1, 1, 22).setValues([updates[i].values]);
  }

  if (appends.length > 0) {
    var start = db.getLastRow() + 1;
    db.getRange(start, 1, appends.length, 22).setValues(appends);
  }

  if (!isSilent) {
    SpreadsheetApp.getUi().alert(
      "✅ DB 동기화 완료\n업데이트 " + updates.length + " / 신규 " + appends.length +
      "\nDB 시트: " + db.getName()
    );
  }
}

/** ✅ 메뉴용: 진행만 */
function syncDBActive() {
  syncProjectDb(false, false);
}

/** ✅ 메뉴용: 전체 스캔 (Main.gs의 syncDBFullScan 에러 해결) */
function syncDBFullScan() {
  syncProjectDb(true, false);
}
