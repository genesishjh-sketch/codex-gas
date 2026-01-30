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
    .addItem('👤 연락처 동기화', 'runContactSync')
    .addSeparator()
    .addItem('🧪 실행 진단', 'runDiagnostics')
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

  for (var r = startRow; r <= lastRow; r += blockHeight) {
    if (stopCtl.check(sheet, r)) break;
    total++;

    var nameVal = sheet.getRange(r + CONFIG.POS_NAME.row, CONFIG.POS_NAME.col).getDisplayValue();
    if (!nameVal) emptyName++;

    if (isValidName(nameVal)) valid++;
    else invalid++;

    if (isClosedBlock_(sheet, r)) closed++;

    var status = (typeof findStatusInRow_ === "function") ? findStatusInRow_(sheet, r) : "";
    if (typeof isActiveStatusForDrive_ === "function" && isActiveStatusForDrive_(status)) active++;
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
  msg.push("- 카카오키 상태: " + (keyOk ? "OK" : "⚠️ 확인 필요"));

  ui.alert(msg.join("\n"));
}
