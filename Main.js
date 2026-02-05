/** Main.gs
 * - 메뉴 복구
 * - 드라이브 체크: 열 때는 “진행만”
 * - 버튼으로 “전체 스캔(완료/취소 포함)” 제공
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('🚀 스마트 통합 관리')
    .addItem('🛠️ 현장 초기 세팅 (주소+폴더+파일+연락처)', 'runMasterSync')
    .addSeparator()
    .addItem('🟡 드라이브 체크 (진행만 / 열 때)', 'runDriveCheckActive')
    .addItem('🟡 드라이브 체크 (전체 스캔)', 'runDriveCheckAll')
    .addSeparator()
    .addItem('📅 금주 일정표 만들기', 'generateWeeklyCalendar')
    .addSeparator()
    .addItem('🔢 잔금일 기준 번호 재정렬', 'renumberByBalanceDate')
    .addSeparator()
    .addItem('👤 연락처 동기화', 'runContactSync')
    .addItem('🔍 연락처 로그 점검', 'runContactAudit')
    .addSeparator()
    .addItem('🧪 오류파악', 'runDiagnostics')
    .addToUi();

  // ✅ “열 때만” 진행 체크 (조용히)
  try {
    runDriveCheckActive(true);
  } catch (e) {}
}

/** 마스터 세팅(팝업 1회) */
function runMasterSync() {
  var ui = SpreadsheetApp.getUi();

  // (주소 변환을 쓰는 경우만 체크)
  if (typeof KAKAO_API_KEY !== "undefined") {
    var k = (KAKAO_API_KEY || "").toString();
    if (!k || k.indexOf("여기에") >= 0) {
      ui.alert("⚠️ 설정 오류\nKAKAO_API_KEY를 확인해주세요!");
      return;
    }
  }

  var addrResult = (typeof updateAddressesBatch === "function") ? updateAddressesBatch(true) : { summary: "주소 변환 함수 없음", failedList: [] };
  var folderResult = (typeof createFoldersBatch === "function") ? createFoldersBatch(true, false) : { summary: "폴더 생성 함수 없음", successList: [], failedList: [] };
  var contactResult = (typeof syncContactsBatch === "function") ? syncContactsBatch(true) : { summary: "연락처 동기화 함수 없음" };

  var finalMsg = "✅ [주소 변환]\n" + (addrResult.summary || "") + "\n";
  if (addrResult.failedList && addrResult.failedList.length > 0) finalMsg += "❌ 실패:\n" + addrResult.failedList.join("\n") + "\n";

  finalMsg += "\n✅ [폴더/파일]\n" + (folderResult.summary || "") + "\n";
  if (folderResult.successList && folderResult.successList.length > 0) finalMsg += "\n✨ 신규 세팅:\n" + folderResult.successList.join("\n") + "\n";
  if (folderResult.failedList && folderResult.failedList.length > 0) finalMsg += "\n❌ 실패:\n" + folderResult.failedList.join("\n") + "\n";

  finalMsg += "\n✅ [연락처]\n" + (contactResult.summary || "") + "\n";

  ui.alert("🎉 작업 완료 리포트\n\n" + finalMsg);

  // 끝나고 진행만 드라이브 체크
  runDriveCheckActive(true);
}

function runContactSync() {
  if (typeof syncContactsBatch !== "function") {
    SpreadsheetApp.getUi().alert("⚠️ syncContactsBatch 함수가 없습니다.");
    return;
  }
  syncContactsBatch(false);
}

function runContactAudit() {
  if (typeof auditContactLog_ !== "function") {
    SpreadsheetApp.getUi().alert("⚠️ auditContactLog_ 함수가 없습니다.");
    return;
  }
  auditContactLog_(false);
}

/** 드라이브 체크: 진행만(열 때) */
function runDriveCheckActive(isSilent) {
  if (typeof checkFolderFilesColor !== "function") {
    SpreadsheetApp.getUi().alert("⚠️ checkFolderFilesColor 함수가 없습니다.");
    return;
  }
  // 캐시 사용(빠름)
  checkFolderFilesColor(!!isSilent, { includeAll: false, forceRefresh: false });
}

/** 드라이브 체크: 전체(완료/취소 포함) */
function runDriveCheckAll() {
  if (typeof checkFolderFilesColor !== "function") {
    SpreadsheetApp.getUi().alert("⚠️ checkFolderFilesColor 함수가 없습니다.");
    return;
  }
  // 강제 재검사(정확)
  checkFolderFilesColor(false, { includeAll: true, forceRefresh: true });
}

/** 실행 진단: 왜 "아무것도 안 됨"인지 빠르게 요약 */
function runDiagnostics() {
  var ui = SpreadsheetApp.getUi();
  var sheet;
  try {
    sheet = getMainSheet_();
  } catch (e) {
    ui.alert("❌ 진단 실패\n시트 오류: " + (e && e.message ? e.message : e));
    return;
  }

  var lastRow = sheet.getLastRow();
  var startRow = getStartRow_();
  var blockHeight = getBlockHeight_(sheet);

  if (lastRow < startRow) {
    ui.alert("ℹ️ 진단 결과\n데이터가 없습니다.\n시작행: " + startRow + " / 마지막행: " + lastRow);
    return;
  }

  var stopCtl = makeStopController_();
  var total = 0, valid = 0, invalid = 0, closed = 0, active = 0, emptyName = 0;
  var invalidSamples = [];
  var statusCounts = {};
  var statusSamples = {};
  var emptyStatus = 0;

  for (var r = startRow; r <= lastRow; r += blockHeight) {
    if (stopCtl.check(sheet, r)) break;
    total++;

    var nameVal = sheet.getRange(r + CONFIG.POS_NAME.row, CONFIG.POS_NAME.col).getDisplayValue();
    if (!nameVal) emptyName++;

    if (isValidName(nameVal)) valid++;
    else invalid++;

    if (isClosedBlock_(sheet, r)) closed++;

    var status = (typeof findStatusInBlock_ === "function") ? findStatusInBlock_(sheet, r) :
      ((typeof findStatusInRow_ === "function") ? findStatusInRow_(sheet, r) : "");
    var statusKey = status || "(없음)";
    statusCounts[statusKey] = (statusCounts[statusKey] || 0) + 1;
    if (!status) emptyStatus++;

    if (!statusSamples[statusKey]) statusSamples[statusKey] = [];
    if (statusSamples[statusKey].length < 3) {
      statusSamples[statusKey].push(r);
    }

    if (typeof isActiveStatusForDrive_ === "function" && isActiveStatusForDrive_(status)) active++;

    if (!isValidName(nameVal) && invalidSamples.length < 5) {
      invalidSamples.push("Row " + r + ": " + nameVal);
    }
  }

  var keyOk = true;
  if (typeof KAKAO_API_KEY !== "undefined") {
    var k = (KAKAO_API_KEY || "").toString();
    if (!k || k.indexOf("여기에") >= 0) keyOk = false;
  }

  var msg = [];
  msg.push("✅ 진단 요약");
  msg.push("- 시트명: " + (CONFIG && CONFIG.SHEET_NAME ? CONFIG.SHEET_NAME : "(미설정)"));
  msg.push("- 시작행/블록높이: " + startRow + " / " + blockHeight);
  msg.push("- 총 블록: " + total);
  msg.push("- 유효 프로젝트명: " + valid + " (무효 " + invalid + ")");
  msg.push("- 프로젝트명 빈칸 블록: " + emptyName);
  msg.push("- 완료/취소 블록: " + closed);
  msg.push("- 진행 상태(드라이브 체크 대상): " + active);
  msg.push("- 상태 없음 블록: " + emptyStatus);
  msg.push("");
  msg.push("✅ 상태 분포(상위)");
  Object.keys(statusCounts).sort(function(a, b){
    return statusCounts[b] - statusCounts[a];
  }).slice(0, 6).forEach(function(key) {
    msg.push("  - " + key + ": " + statusCounts[key] + " (예: " + statusSamples[key].join(", ") + ")");
  });

  if (invalidSamples.length > 0) {
    msg.push("");
    msg.push("⚠️ 무효 프로젝트명 예시(최대 5개)");
    invalidSamples.forEach(function(v) { msg.push("  - " + v); });
  }
  msg.push("- 카카오키 상태: " + (keyOk ? "OK" : "⚠️ 확인 필요"));

  ui.alert(msg.join("\n"));
}

/** 잔금일(예정/완료) 날짜순으로 B열 번호 재정렬 */
function renumberByBalanceDate() {
  var ui = SpreadsheetApp.getUi();
  var sheet;
  try {
    sheet = getMainSheet_();
  } catch (e) {
    ui.alert("❌ 시트 오류\n" + (e && e.message ? e.message : e));
    return;
  }

  var blockHeight = getBlockHeight_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.START_ROW) {
    ui.alert("ℹ️ 데이터가 없습니다.");
    return;
  }

  var stopCtl = makeStopController_();
  var blocks = [];
  var datedBlocks = [];

  function isEmpty_(val) {
    if (val === null || val === undefined) return true;
    if (val instanceof Date) return false;
    return String(val).trim() === "";
  }
  function toDateKey_(val) {
    if (val instanceof Date) return val.getTime();
    return null;
  }

  for (var r = CONFIG.START_ROW; r <= lastRow; r += blockHeight) {
    if (stopCtl.check(sheet, r)) break;

    var d5 = sheet.getRange(r + 1, 4).getValue(); // D5
    var d9 = sheet.getRange(r + 5, 4).getValue(); // D9
    var dateKey = toDateKey_(d9);
    var hasD5 = !isEmpty_(d5);
    var hasD9 = !isEmpty_(d9);

    var entry = {
      row: r,
      d5: d5,
      d9: d9,
      dateKey: dateKey,
      hasD5: hasD5,
      hasD9: hasD9
    };
    blocks.push(entry);
    if (dateKey !== null) datedBlocks.push(entry);
  }

  if (blocks.length === 0) {
    ui.alert("ℹ️ 정렬할 프로젝트가 없습니다.");
    return;
  }

  datedBlocks.sort(function(a, b) {
    if (a.dateKey !== b.dateKey) return a.dateKey - b.dateKey;
    return a.row - b.row;
  });

  var sequenceMap = {};
  for (var i = 0; i < datedBlocks.length; i++) {
    sequenceMap[datedBlocks[i].row] = String(i + 1).padStart(3, "0");
  }

  blocks.forEach(function(block) {
    var value = "";
    if (block.hasD9) {
      value = sequenceMap[block.row] || "";
    } else if (block.hasD5) {
      value = "999";
    } else {
      value = "";
    }
    sheet.getRange(block.row, CONFIG.POS_NO.col).setValue(value);
  });

  try {
    sortGroupsByHierarchy();
  } catch (e) {
    ui.alert("❌ 그룹 정렬 중 오류가 발생했습니다.\n" + (e && e.message ? e.message : e));
    return;
  }

  ui.alert("✅ 잔금일 기준 번호 재정렬 완료\n대상: " + blocks.length + "건 / 잔금일: " + datedBlocks.length + "건");
}

/** 그룹(9행) 단위로 우선순위에 따라 정렬 */
function sortGroupsByHierarchy() {
  var sheet = getMainSheet_();
  var blockHeight = getBlockHeight_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.START_ROW) return;

  var lastCol = sheet.getLastColumn();
  var groups = [];

  function isEmpty_(val) {
    if (val === null || val === undefined) return true;
    if (val instanceof Date) return false;
    return String(val).trim() === "";
  }

  function compareValues_(a, b) {
    if (a instanceof Date && b instanceof Date) return a.getTime() - b.getTime();
    if (typeof a === "number" && typeof b === "number") return a - b;
    return String(a).localeCompare(String(b));
  }

  for (var r = CONFIG.START_ROW; r <= lastRow; r += blockHeight) {
    var bVal = sheet.getRange(r, 2).getValue();
    var cVal = sheet.getRange(r + 1, 3).getValue();
    var hasB = !isEmpty_(bVal);
    var hasC = !isEmpty_(cVal);
    var priority = 3;
    if (hasB) {
      priority = 1;
    } else if (hasC) {
      priority = 2;
    }
    groups.push({
      startRow: r,
      bVal: bVal,
      priority: priority,
      originalIndex: groups.length
    });
  }

  if (groups.length === 0) return;

  var sortedGroups = groups.slice().sort(function(a, b) {
    if (a.priority !== b.priority) return a.priority - b.priority;
    if (a.priority === 1) {
      var cmp = compareValues_(a.bVal, b.bVal);
      if (cmp !== 0) return cmp;
    }
    return a.originalIndex - b.originalIndex;
  });

  var needsMove = sortedGroups.some(function(group, index) {
    return group.startRow !== CONFIG.START_ROW + index * blockHeight;
  });
  if (!needsMove) return;

  // 원본을 임시 영역으로 복사한 뒤 정렬 순서대로 다시 복사한다.
  var totalRows = groups.length * blockHeight;
  var tempStartRow = lastRow + 1;
  sheet.insertRowsAfter(lastRow, totalRows);

  var tempRow = tempStartRow;
  groups.forEach(function(group) {
    var source = sheet.getRange(group.startRow, 1, blockHeight, lastCol);
    var target = sheet.getRange(tempRow, 1, blockHeight, lastCol);
    source.copyTo(target, { contentsOnly: false });
    tempRow += blockHeight;
  });

  var targetRow = CONFIG.START_ROW;
  sortedGroups.forEach(function(group) {
    var sourceRow = tempStartRow + group.originalIndex * blockHeight;
    var source = sheet.getRange(sourceRow, 1, blockHeight, lastCol);
    var target = sheet.getRange(targetRow, 1, blockHeight, lastCol);
    source.copyTo(target, { contentsOnly: false });
    targetRow += blockHeight;
  });

  sheet.deleteRows(tempStartRow, totalRows);
}
