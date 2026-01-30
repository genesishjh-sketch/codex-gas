/** ReportManager.gs
 *
 * ✅ 드라이브 체크:
 *    - “진행만(열 때)” / “전체 스캔(버튼)” 분리
 *    - 표시(색칠)는 폴더 URL 셀(S)만 (R은 폴더명)
 *    - 메인폴더 + 1~2단 서브폴더까지 파일 존재 여부 확인
 *    - 오류(권한/쿼터) 시 기존 배경색 유지 → “노란색 풀림” 방지
 */

function checkFolderFilesColor(isSilent, options) {
  options = options || {};
  var includeAll = !!options.includeAll;         // true면 완료/취소/대기 포함
  var forceRefresh = !!options.forceRefresh;     // true면 캐시 무시하고 강제 재검사
  return driveCheckUpdate_(includeAll, forceRefresh, !!isSilent);
}

/** === 내부: 드라이브 체크 본체 === */
function driveCheckUpdate_(includeAll, forceRefresh, isSilent) {
  var sheet = getMainSheet_();
  var blockHeight = getBlockHeight_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < getStartRow_()) return;

  var stopCtl = makeStopController_();

  // (옵션) 캐시 시트 사용: 없으면 자동 생성
  var logName = (CONFIG && CONFIG.DRIVE_LOG_SHEET) ? CONFIG.DRIVE_LOG_SHEET : "드라이브_check_log";
  var logSh = ensureSheet_(logName, ["FOLDER_ID", "HAS_FILES", "LAST_CHECK_AT"]);
  var cacheMap = readDriveCacheMap_(logSh);

  var runLogSh = ensureSheet_("드라이브_run_log", ["RUN_ID","TIME","BLOCK_ROW","STATUS","URL","RESULT","DETAILS"]);
  var runId = Utilities.getUuid();
  var logRows = [];

  var updated = 0, skipped = 0, errors = 0;

  for (var r = getStartRow_(); r <= lastRow; r += blockHeight) {
    if (stopCtl.check(sheet, r)) break;

    // 프로젝트 유효성
    var pname = findProjectNameInRow_(sheet, r);
    if (!isValidName(pname)) continue;

    // 상태 판별(행 전체에서 키워드 검색)
    var status = findStatusInRow_(sheet, r);

    // “열 때(진행만)” 모드면 진행 단계만 검사
    if (!includeAll) {
      if (!isActiveStatusForDrive_(status)) {
        skipped++;
        logRows.push([runId, new Date(), r, status || "", "", "SKIP_STATUS", "진행만 대상 아님"]);
        continue;
      }
    }

    // 폴더 URL 셀 찾기: 블록 내 R(폴더명) 행의 S 링크 우선, 없으면 S열에서 첫 drive URL
    var folderCellInfo = findFolderUrlCell_(sheet, r);
    if (!folderCellInfo) {
      skipped++;
      logRows.push([runId, new Date(), r, status || "", "", "SKIP_NO_URL", "R/S에서 링크를 찾지 못함"]);
      continue;
    }

    var urlCell = folderCellInfo.cell;    // 색칠 대상(=S)
    var folderUrl = folderCellInfo.url;

    // 에러 시 “기존색 유지” 위해 현재 배경 확보
    var prevBg = urlCell.getBackground();

    // URL 없으면 흰색 처리
    if (!folderUrl || folderUrl.indexOf("drive.google.com") === -1) {
      urlCell.setBackground("#ffffff");
      logRows.push([runId, new Date(), r, status || "", folderUrl || "", "SKIP_BAD_URL", "drive URL 아님"]);
      continue;
    }

    var folderId = extractIdFromUrl(folderUrl);
    if (!folderId) {
      urlCell.setBackground("#ffffff");
      logRows.push([runId, new Date(), r, status || "", folderUrl || "", "SKIP_NO_ID", "폴더 ID 추출 실패"]);
      continue;
    }

    try {
      var cached = cacheMap[folderId];
      var hasFiles;
      var source = "SCAN";

      // 전체스캔(forceRefresh)면 캐시 무시
      if (!forceRefresh && cached && isCacheFresh_(cached.at)) {
        hasFiles = !!cached.has;
        source = "CACHE";
      } else {
        // 메인+서브폴더까지 확인 (depth=2)
        hasFiles = folderHasAnyFilesDeep_(folderId, 2);

        // 캐시 기록
        if (cached) {
          logSh.getRange(cached.row, 2, 1, 2).setValues([[hasFiles, new Date()]]);
        } else {
          logSh.appendRow([folderId, hasFiles, new Date()]);
        }
      }

      logRows.push([runId, new Date(), r, status || "", folderUrl, "UPDATED", (hasFiles ? "HAS_FILES" : "NO_FILES") + " / " + source]);
      urlCell.setBackground(hasFiles ? "#ffff00" : "#ffffff");
      updated++;

    } catch (e) {
      // 권한/쿼터/일시 오류 → 기존색 유지
      urlCell.setBackground(prevBg);
      errors++;
      logRows.push([runId, new Date(), r, status || "", folderUrl || "", "ERROR", String(e && e.message ? e.message : e)]);
      continue;
    }
  }

  if (logRows.length > 0) appendLogRows_(runLogSh, logRows);

  if (!isSilent) {
    SpreadsheetApp.getUi().alert(
      "✅ 드라이브 체크 완료\n업데이트 " + updated + " / 스킵 " + skipped + " / 오류 " + errors
    );
  }

  return { updated: updated, skipped: skipped, errors: errors };
}

/** ===== 로그/캐시 ===== */

function appendLogRows_(sheet, rows) {
  if (!rows || rows.length === 0) return;
  var start = sheet.getLastRow() + 1;
  sheet.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
}

function readDriveCacheMap_(logSh) {
  var last = logSh.getLastRow();
  var map = {};
  if (last < 2) return map;

  var vals = logSh.getRange(2, 1, last - 1, 3).getValues();
  for (var i = 0; i < vals.length; i++) {
    var id = (vals[i][0] || "").toString().trim();
    if (!id) continue;
    map[id] = { row: i + 2, has: !!vals[i][1], at: vals[i][2] };
  }
  return map;
}

function isCacheFresh_(dt) {
  if (!(dt instanceof Date)) return false;
  var hours = (typeof CONFIG !== "undefined" && CONFIG.DRIVE_CACHE_HOURS) ? CONFIG.DRIVE_CACHE_HOURS : 6;
  return (new Date().getTime() - dt.getTime()) <= (hours * 3600 * 1000);
}

/**
 * 폴더 내 실제 파일 존재 여부
 * - 메인 폴더 파일 검사
 * - 없으면 서브폴더(최대 depth)까지 검사
 * - '물품리스트'는 제외
 */
function folderHasAnyFilesDeep_(folderId, depth) {
  var folder = DriveApp.getFolderById(folderId);

  // 1) 메인 폴더 파일
  if (folderHasActualFiles_(folder)) return true;

  // 2) 서브폴더 탐색
  if (!depth || depth <= 0) return false;

  var sub = folder.getFolders();
  while (sub.hasNext()) {
    var sf = sub.next();
    if (folderHasActualFiles_(sf)) return true;
    if (depth > 1) {
      // 2단계까지
      var sub2 = sf.getFolders();
      while (sub2.hasNext()) {
        var sf2 = sub2.next();
        if (folderHasActualFiles_(sf2)) return true;
      }
    }
  }
  return false;
}

function folderHasActualFiles_(folder) {
  var it = folder.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    var nm = (f.getName() || "");
    if (nm.indexOf("물품리스트") >= 0) continue;
    return true;
  }
  return false;
}

/** ===== 시트 구조 탐색 ===== */

/**
 * 블록 내 R(폴더명) 행의 S 링크를 우선 사용하고,
 * 없으면 S열에서 첫 번째 drive URL을 찾아 반환.
 * (색칠은 항상 S(URL) 셀만)
 */
function findFolderUrlCell_(sheet, blockStartRow) {
  var lastRow = sheet.getLastRow();
  var blockHeight = getBlockHeight_(sheet);
  var endRow = Math.min(lastRow, blockStartRow + blockHeight - 1);

  // Prefer fixed label/url columns (R/S). Use label to select the row, then read S link.
  var labelCol = (CONFIG && CONFIG.POS_FOLDER_LABEL_COL) || 18; // R
  var urlCol = (CONFIG && (CONFIG.POS_FOLDER_URL_COL || CONFIG.DRIVE_MARK_COL)) || 19; // S
  var scanRows = endRow - blockStartRow + 1;

  if (labelCol && urlCol && scanRows > 0) {
    var labelRange = sheet.getRange(blockStartRow, labelCol, scanRows, 1);
    var urlRange = sheet.getRange(blockStartRow, urlCol, scanRows, 1);
    var labelVals = labelRange.getDisplayValues();
    var urlVals = urlRange.getDisplayValues();

    for (var i = 0; i < labelVals.length; i++) {
      var label = (labelVals[i][0] || "").toString().trim();
      if (!label) continue;

      var urlCell = urlRange.getCell(i + 1, 1);
      var url = (urlVals[i][0] || "").toString().trim();
      if (!url) {
        url = getUrlFromCell_(urlCell);
      }
      if (url && url.indexOf("drive.google.com") >= 0) {
        return { cell: urlCell, url: url, col: urlCol };
      }
    }
  }

  // If no labeled row matched, scan S column in the block for any drive URL.
  if (urlCol && scanRows > 0) {
    var urlRange2 = sheet.getRange(blockStartRow, urlCol, scanRows, 1);
    var urlVals2 = urlRange2.getDisplayValues();
    for (var j = 0; j < urlVals2.length; j++) {
      var cell2 = urlRange2.getCell(j + 1, 1);
      var url2 = (urlVals2[j][0] || "").toString().trim();
      if (!url2) {
        url2 = getUrlFromCell_(cell2);
      }
      if (url2 && url2.indexOf("drive.google.com") >= 0) {
        return { cell: cell2, url: url2, col: urlCol };
      }
    }
  }

  // Fallback: label-based scan on the block start row (legacy layout)
  var maxCols = Math.min(sheet.getLastColumn(), 220);
  var rowVals = sheet.getRange(blockStartRow, 1, 1, maxCols).getDisplayValues()[0];
  for (var c = 0; c < rowVals.length; c++) {
    var v = (rowVals[c] || "").toString().trim();
    if (v === "[폴더]" && c + 1 < rowVals.length) {
      var urlCol2 = c + 2; // 1-based
      var cell3 = sheet.getRange(blockStartRow, urlCol2);
      var url3 = cell3.getDisplayValue().toString().trim();
      if (!url3) {
        var rich = cell3.getRichTextValue();
        if (rich && rich.getLinkUrl()) url3 = rich.getLinkUrl();
      }
      return { cell: cell3, url: url3, col: urlCol2 };
    }
  }

  // Fallback: search for a drive URL in the block start row
  for (var k = 0; k < rowVals.length; k++) {
    var s = (rowVals[k] || "").toString();
    if (s.indexOf("drive.google.com") >= 0) {
      var cell4 = sheet.getRange(blockStartRow, k + 1);
      return { cell: cell4, url: s.trim(), col: k + 1 };
    }
  }
  return null;
}

function findStatusInRow_(sheet, blockStartRow) {
  var maxCols = Math.min(sheet.getLastColumn(), 220);
  var vals = sheet.getRange(blockStartRow, 1, 1, maxCols).getDisplayValues()[0];
  var candidates = [
    "취소",
    "완료",
    "세팅완료(에비대기)",
    "대기",
    "세팅 대기",
    "엑셀 작업",
    "디자인 작업",
    "실측대기",
    "진행"
  ];
  for (var i = 0; i < candidates.length; i++) {
    for (var c = 0; c < vals.length; c++) {
      if ((vals[c] || "").toString().trim() === candidates[i]) return candidates[i];
    }
  }
  return "";
}

function isActiveStatusForDrive_(status) {
  // “열 때(진행만)”에서 체크할 상태들
  return (
    status === "진행" ||
    status === "실측대기" ||
    status === "디자인 작업" ||
    status === "엑셀 작업" ||
    status === "세팅 대기"
  );
}

function findProjectNameInRow_(sheet, blockStartRow) {
  var maxCols = Math.min(sheet.getLastColumn(), 40);
  var vals = sheet.getRange(blockStartRow, 1, 1, maxCols).getDisplayValues()[0];
  for (var i = 0; i < vals.length; i++) {
    var s = (vals[i] || "").toString().trim();
    if (!s) continue;
    // 프로젝트명 패턴(멱살반/반멱살/스타일링대행 등) 우선
    if (s.indexOf("멱살") >= 0 || s.indexOf("스타일") >= 0) return s;
  }
  // fallback: 첫 번째 non-empty
  for (var j = 0; j < vals.length; j++) {
    var t = (vals[j] || "").toString().trim();
    if (t) return t;
  }
  return "";
}
