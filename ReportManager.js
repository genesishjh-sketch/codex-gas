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

function getDriveCheckSheets_(includeAll) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configuredNames = includeAll && CONFIG && CONFIG.DRIVE_CHECK_SHEET_NAMES && CONFIG.DRIVE_CHECK_SHEET_NAMES.length
    ? CONFIG.DRIVE_CHECK_SHEET_NAMES
    : [CONFIG.SHEET_NAME, "완료"];
  if (!includeAll) configuredNames = [CONFIG.SHEET_NAME];
  var seen = {};
  var sheets = [];

  for (var i = 0; i < configuredNames.length; i++) {
    var name = String(configuredNames[i] || "").trim();
    if (!name || seen[name]) continue;
    seen[name] = true;

    var sheet = ss.getSheetByName(name);
    if (sheet) sheets.push(sheet);
  }

  if (sheets.length === 0) sheets.push(getMainSheet_());
  return sheets;
}

function inspectDriveAndContacts(isSilent, options) {
  options = options || {};
  var includeAll = !!options.includeAll;
  var forceRefresh = !!options.forceRefresh;
  var sheets = getDriveCheckSheets_(includeAll);
  var reportHeaders = [
    "CHECKED_AT", "SHEET_NAME", "BLOCK_ROW", "PROJECT_NAME", "STATUS", "PHONE",
    "CONTACT_EXISTS", "DRIVE_HAS_FILES", "DRIVE_URL", "DETAILS"
  ];

  var reportSheet = ensureSheet_("진짜있나", reportHeaders);
  reportSheet.getRange(1, 1, 1, reportHeaders.length).setValues([reportHeaders]).setFontWeight("bold");
  reportSheet.setFrozenRows(1);
  if (reportSheet.getLastRow() > 1) {
    reportSheet.getRange(2, 1, reportSheet.getLastRow() - 1, reportHeaders.length).clearContent();
  }

  var logName = (CONFIG && CONFIG.DRIVE_LOG_SHEET) ? CONFIG.DRIVE_LOG_SHEET : "드라이브_check_log";
  var logSh = ensureSheet_(logName, ["FOLDER_ID", "HAS_FILES", "LAST_CHECK_AT"]);
  var cacheMap = readDriveCacheMap_(logSh);
  var runLogSh = ensureSheet_("드라이브_run_log", ["RUN_ID","TIME","BLOCK_ROW","STATUS","URL","RESULT","DETAILS"]);
  var runId = Utilities.getUuid();
  var driveLogRows = [];

  var contactsState = (typeof getContactsServiceState_ === "function")
    ? getContactsServiceState_()
    : { ok: false, reason: "contact_checker_unavailable" };

  var rows = [];
  var checked = 0;
  var skipped = 0;
  var driveErrors = 0;
  var contactErrors = 0;

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var sheetName = sheet.getName();
    var blockHeight = getBlockHeight_(sheet);
    var lastRow = sheet.getLastRow();
    if (lastRow < getStartRow_()) continue;

    var stopCtl = makeStopController_();
    for (var r = getStartRow_(); r <= lastRow; r += blockHeight) {
      if (stopCtl.check(sheet, r)) break;

      var projectName = findProjectNameInRow_(sheet, r);
      if (!isValidName(projectName)) continue;

      var status = findStatusInBlock_(sheet, r);
      if (!includeAll && isClosedBlock_(sheet, r)) {
        skipped++;
        continue;
      }

      var contactInfo = (typeof getContactBlockInfo_ === "function")
        ? getContactBlockInfo_(sheet, r, blockHeight, projectName)
        : { phone: "", normalized: "" };
      var phone = contactInfo.normalized || contactInfo.phone || "";
      var contactResult = inspectContactExists_(phone, contactsState);
      if (contactResult.error) contactErrors++;

      var driveResult = inspectDriveForBlock_(sheet, r, status, forceRefresh, cacheMap, logSh, driveLogRows, runId, sheetName);
      if (driveResult.error) driveErrors++;

      rows.push([
        new Date(),
        sheetName,
        r,
        projectName,
        status || "",
        phone,
        contactResult.text,
        driveResult.text,
        driveResult.url || "",
        [contactResult.detail, driveResult.detail].filter(function(v) { return !!v; }).join(" / ")
      ]);
      checked++;
    }
  }

  if (rows.length > 0) {
    reportSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    reportSheet.autoResizeColumns(1, reportHeaders.length);
  }
  if (driveLogRows.length > 0) appendLogRows_(runLogSh, driveLogRows);

  if (!isSilent) {
    SpreadsheetApp.getUi().alert(
      "✅ 드라이브 및 연락처 검사 완료\n" +
      "검사 " + checked + " / 제외 " + skipped +
      " / 드라이브 오류 " + driveErrors +
      " / 연락처 오류 " + contactErrors +
      "\n결과 시트: 진짜있나"
    );
  }

  return { checked: checked, skipped: skipped, driveErrors: driveErrors, contactErrors: contactErrors };
}

function inspectContactExists_(phone, contactsState) {
  if (!phone) return { text: "NO_PHONE", detail: "연락처 없음" };
  if (!contactsState || !contactsState.ok) {
    return {
      text: "CHECK_UNAVAILABLE",
      detail: "연락처 검사 불가: " + ((contactsState && contactsState.reason) || "unknown"),
      error: true
    };
  }
  if (typeof findContactsByPhone_ !== "function") {
    return { text: "CHECK_UNAVAILABLE", detail: "findContactsByPhone_ 없음", error: true };
  }

  try {
    var found = findContactsByPhone_(phone);
    return {
      text: (found && found.length > 0) ? "YES" : "NO",
      detail: (found && found.length > 0) ? "연락처 있음" : "연락처 없음"
    };
  } catch (e) {
    return {
      text: "ERROR",
      detail: "연락처 검사 오류: " + (e && e.message ? e.message : e),
      error: true
    };
  }
}

function inspectDriveForBlock_(sheet, blockStartRow, status, forceRefresh, cacheMap, logSh, driveLogRows, runId, sheetName) {
  sheetName = sheetName || sheet.getName();
  var withSheet = function(detail) {
    return sheetName + " / " + detail;
  };
  var folderCellInfos = findFolderUrlCells_(sheet, blockStartRow);
  if (!folderCellInfos || folderCellInfos.length === 0) {
    driveLogRows.push([runId, new Date(), blockStartRow, status || "", "", "SKIP_NO_URL", withSheet("R/S에서 링크를 찾지 못함")]);
    return { text: "NO_URL", detail: "드라이브 URL 없음", url: "" };
  }

  var urls = [];
  var hasAnyFiles = false;
  var usedCache = false;
  var skippedYellow = false;
  var errors = [];

  for (var i = 0; i < folderCellInfos.length; i++) {
    var folderCellInfo = folderCellInfos[i];
    var urlCell = folderCellInfo.cell;
    var folderUrl = folderCellInfo.url || "";
    var prevBg = urlCell.getBackground();
    urls.push(folderUrl);

    if (!forceRefresh && isMarkedHasFilesBackground_(prevBg)) {
      hasAnyFiles = true;
      skippedYellow = true;
      driveLogRows.push([runId, new Date(), blockStartRow, status || "", folderUrl, "SKIP_YELLOW", withSheet("노란색 표시 유지")]);
      continue;
    }

    if (!folderUrl || folderUrl.indexOf("drive.google.com") === -1) {
      urlCell.setBackground("#ffffff");
      driveLogRows.push([runId, new Date(), blockStartRow, status || "", folderUrl, "SKIP_BAD_URL", withSheet("drive URL 아님")]);
      continue;
    }

    var folderId = extractIdFromUrl(folderUrl);
    if (!folderId || folderId.indexOf("http") >= 0 || folderId.indexOf("/") >= 0) {
      urlCell.setBackground("#ffffff");
      driveLogRows.push([runId, new Date(), blockStartRow, status || "", folderUrl, "SKIP_NO_ID", withSheet("폴더 ID 추출 실패")]);
      continue;
    }

    try {
      var cached = cacheMap[folderId];
      var hasFiles;
      var source = "SCAN";
      if (!forceRefresh && cached && isCacheFresh_(cached.at)) {
        hasFiles = !!cached.has;
        source = "CACHE";
        usedCache = true;
      } else {
        hasFiles = folderHasAnyFilesDeep_(folderId, 2);
        if (cached) {
          logSh.getRange(cached.row, 2, 1, 2).setValues([[hasFiles, new Date()]]);
          cached.has = hasFiles;
          cached.at = new Date();
        } else {
          logSh.appendRow([folderId, hasFiles, new Date()]);
          cacheMap[folderId] = { row: logSh.getLastRow(), has: hasFiles, at: new Date() };
        }
      }

      if (hasFiles) hasAnyFiles = true;
      urlCell.setBackground(hasFiles ? "#ffff00" : "#ffffff");
      driveLogRows.push([runId, new Date(), blockStartRow, status || "", folderUrl, "UPDATED", withSheet((hasFiles ? "HAS_FILES" : "NO_FILES") + " / " + source)]);
    } catch (e) {
      urlCell.setBackground(prevBg);
      errors.push(e && e.message ? e.message : String(e));
      driveLogRows.push([runId, new Date(), blockStartRow, status || "", folderUrl, "ERROR", withSheet(String(e && e.message ? e.message : e))]);
    }
  }

  if (errors.length > 0) {
    return {
      text: hasAnyFiles ? "HAS_FILES_WITH_ERROR" : "ERROR",
      detail: errors.join(" | "),
      url: urls.filter(function(v) { return !!v; }).join("\n"),
      error: true
    };
  }

  return {
    text: hasAnyFiles ? "YES" : "NO",
    detail: skippedYellow ? "노란색 셀 검사 생략" : (usedCache ? "캐시 사용" : "새로 검사"),
    url: urls.filter(function(v) { return !!v; }).join("\n")
  };
}

function isMarkedHasFilesBackground_(color) {
  var normalized = String(color || "").trim().toLowerCase();
  return normalized === "#ffff00" || normalized === "yellow";
}

/** === 내부: 드라이브 체크 본체 === */
function driveCheckUpdate_(includeAll, forceRefresh, isSilent) {
  if (typeof DriveApp === "undefined") {
    if (!isSilent) SpreadsheetApp.getUi().alert("⚠️ DriveApp을 사용할 수 없어 드라이브 체크를 중단합니다.");
    return { updated: 0, skipped: 0, errors: 0 };
  }
  try {
    DriveApp.getRootFolder();
  } catch (e) {
    if (!isSilent) {
      SpreadsheetApp.getUi().alert("⚠️ DriveApp 권한이 없어 드라이브 체크를 중단합니다.\n" + (e && e.message ? e.message : e));
    }
    return { updated: 0, skipped: 0, errors: 1 };
  }
  var sheets = getDriveCheckSheets_(includeAll);

  // (옵션) 캐시 시트 사용: 없으면 자동 생성
  var logName = (CONFIG && CONFIG.DRIVE_LOG_SHEET) ? CONFIG.DRIVE_LOG_SHEET : "드라이브_check_log";
  var logSh = ensureSheet_(logName, ["FOLDER_ID", "HAS_FILES", "LAST_CHECK_AT"]);
  var cacheMap = readDriveCacheMap_(logSh);

  var runLogSh = ensureSheet_("드라이브_run_log", ["RUN_ID","TIME","BLOCK_ROW","STATUS","URL","RESULT","DETAILS"]);
  var runId = Utilities.getUuid();
  var logRows = [];

  var updated = 0, skipped = 0, errors = 0;

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var sheetName = sheet.getName();
    var blockHeight = getBlockHeight_(sheet);
    var lastRow = sheet.getLastRow();
    if (lastRow < getStartRow_()) continue;

    var stopCtl = makeStopController_();
    for (var r = getStartRow_(); r <= lastRow; r += blockHeight) {
      if (stopCtl.check(sheet, r)) break;

      // 프로젝트 유효성
      var pname = findProjectNameInRow_(sheet, r);
      if (!isValidName(pname)) continue;

      // 상태 판별(블록 전체에서 키워드 검색)
      var status = findStatusInBlock_(sheet, r);

      // “열 때(진행만)” 모드면 진행 단계만 검사
      if (!includeAll) {
        if (!isActiveStatusForDrive_(status)) {
          skipped++;
          logRows.push([runId, new Date(), r, status || "", "", "SKIP_STATUS", sheetName + " / 진행만 대상 아님"]);
          continue;
        }
      }

      // 폴더 URL 셀 찾기: 블록 내 S열 링크 전체 수집
      var folderCellInfos = findFolderUrlCells_(sheet, r);
      if (!folderCellInfos || folderCellInfos.length === 0) {
        skipped++;
        logRows.push([runId, new Date(), r, status || "", "", "SKIP_NO_URL", sheetName + " / R/S에서 링크를 찾지 못함"]);
        continue;
      }

      for (var f = 0; f < folderCellInfos.length; f++) {
        var folderCellInfo = folderCellInfos[f];
        var urlCell = folderCellInfo.cell;    // 색칠 대상(=S)
        var folderUrl = folderCellInfo.url;

        // 에러 시 “기존색 유지” 위해 현재 배경 확보
        var prevBg = urlCell.getBackground();

        if (!forceRefresh && isMarkedHasFilesBackground_(prevBg)) {
          skipped++;
          logRows.push([runId, new Date(), r, status || "", folderUrl || "", "SKIP_YELLOW", sheetName + " / 노란색 표시 유지"]);
          continue;
        }

        // URL 없으면 흰색 처리
        if (!folderUrl || folderUrl.indexOf("drive.google.com") === -1) {
          urlCell.setBackground("#ffffff");
          logRows.push([runId, new Date(), r, status || "", folderUrl || "", "SKIP_BAD_URL", sheetName + " / drive URL 아님"]);
          continue;
        }

        var folderId = extractIdFromUrl(folderUrl);
        if (!folderId || folderId.indexOf("http") >= 0 || folderId.indexOf("/") >= 0) {
          urlCell.setBackground("#ffffff");
          logRows.push([runId, new Date(), r, status || "", folderUrl || "", "SKIP_NO_ID", sheetName + " / 폴더 ID 추출 실패"]);
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
              cached.has = hasFiles;
              cached.at = new Date();
            } else {
              logSh.appendRow([folderId, hasFiles, new Date()]);
              cacheMap[folderId] = { row: logSh.getLastRow(), has: hasFiles, at: new Date() };
            }
          }

          logRows.push([runId, new Date(), r, status || "", folderUrl, "UPDATED", sheetName + " / " + (hasFiles ? "HAS_FILES" : "NO_FILES") + " / " + source]);
          urlCell.setBackground(hasFiles ? "#ffff00" : "#ffffff");
          updated++;

        } catch (e) {
          // 권한/쿼터/일시 오류 → 기존색 유지
          urlCell.setBackground(prevBg);
          errors++;
          logRows.push([runId, new Date(), r, status || "", folderUrl || "", "ERROR", sheetName + " / " + String(e && e.message ? e.message : e)]);
          continue;
        }
      }
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
function findFolderUrlCells_(sheet, blockStartRow) {
  var lastRow = sheet.getLastRow();
  var blockHeight = getBlockHeight_(sheet);
  var endRow = Math.min(lastRow, blockStartRow + blockHeight - 1);

  // Prefer fixed label/url columns (R/S). Use label to select the row, then read S link.
  var labelCol = (CONFIG && CONFIG.POS_FOLDER_LABEL_COL) || 18; // R
  var urlCol = (CONFIG && (CONFIG.POS_FOLDER_URL_COL || CONFIG.DRIVE_MARK_COL)) || 19; // S
  var scanRows = endRow - blockStartRow + 1;

  var results = [];
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
        results.push({ cell: urlCell, url: url, col: urlCol });
      }
    }
  }

  // Scan S column in the block for any drive URL.
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
        results.push({ cell: cell2, url: url2, col: urlCol });
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
      results.push({ cell: cell3, url: url3, col: urlCol2 });
    }
  }

  // Fallback: search for a drive URL in the block start row
  for (var k = 0; k < rowVals.length; k++) {
    var s = (rowVals[k] || "").toString();
    if (s.indexOf("drive.google.com") >= 0) {
      var cell4 = sheet.getRange(blockStartRow, k + 1);
      results.push({ cell: cell4, url: s.trim(), col: k + 1 });
    }
  }
  if (!results.length) return [];
  return dedupeFolderUrlCells_(results);
}

function dedupeFolderUrlCells_(items) {
  var seen = {};
  var out = [];
  for (var i = 0; i < items.length; i++) {
    var cell = items[i].cell;
    var key = cell.getRow() + ":" + cell.getColumn();
    if (seen[key]) continue;
    seen[key] = true;
    out.push(items[i]);
  }
  return out;
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
    "세팅대기",
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

function findStatusInBlock_(sheet, blockStartRow) {
  var maxCols = Math.min(sheet.getLastColumn(), 220);
  var blockHeight = getBlockHeight_(sheet);
  var lastRow = sheet.getLastRow();
  var endRow = Math.min(lastRow, blockStartRow + blockHeight - 1);
  if (endRow < blockStartRow) return "";

  var vals = sheet.getRange(blockStartRow, 1, endRow - blockStartRow + 1, maxCols).getDisplayValues();
  var candidates = [
    "취소",
    "완료",
    "세팅완료(에비대기)",
    "대기",
    "세팅 대기",
    "세팅대기",
    "엑셀 작업",
    "디자인 작업",
    "실측대기",
    "진행"
  ];

  for (var i = 0; i < candidates.length; i++) {
    var target = candidates[i];
    for (var r = 0; r < vals.length; r++) {
      for (var c = 0; c < vals[r].length; c++) {
        if ((vals[r][c] || "").toString().trim() === target) return target;
      }
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
    status === "세팅 대기" ||
    status === "세팅대기"
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
