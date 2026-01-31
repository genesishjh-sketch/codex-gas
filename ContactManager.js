/** ContactManager.gs (ContactsApp 기반 + 연락처_log 캐시)
 *
 * ✅ 전화번호 위치 고정:
 *   - D6, D15, D24 ... = 각 블록 시작행 + 2행, D열(4)
 */

function getContactLogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = CONFIG.CONTACT_LOG_SHEET_NAME || "연락처_log";
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 9).setValues([[
      "PHONE", "NAME", "PROJECT", "ADDRESS", "MAP_URL",
      "RESULT", "CONTACT_EXISTED", "SOURCE_ROW", "SYNC_AT"
    ]]).setFontWeight("bold");
  }
  return sh;
}

function loadContactLogMap_(logSheet) {
  var map = {};
  var last = logSheet.getLastRow();
  if (last < 2) return map;

  var phones = logSheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < phones.length; i++) {
    var p = (phones[i][0] || "").toString().trim();
    if (p) map[p] = i + 2;
  }
  return map;
}

/**
 * ✅ 전화번호는 고정셀(D열, 블록시작+2)만 읽는다
 * - D6, D15, D24 ...
 */
function findPhoneInBlock_(sheet, blockStartRow, blockHeight) {
  var phoneCellRowOffset = 2; // blockStartRow + 2 => D6 패턴
  var phoneCol = 4;           // D열

  var raw = sheet.getRange(blockStartRow + phoneCellRowOffset, phoneCol).getDisplayValue();
  var s = (raw || "").toString();

  // 010-0000-0000 / 010 0000 0000 / (010)0000-0000 / 01000000000
  var phoneRegex = /\(?01[016789]\)?[\s\-.]?\d{3,4}[\s\-.]?\d{4}/;

  var m = s.match(phoneRegex);
  return (m && m[0]) ? normalizePhone_(m[0]) : "";
}

function ensureContact_(displayName, phone, addressLine, mapUrl) {
  if (!phone) return { ok: false, skipped: true, reason: "no_phone" };

  var normalized = normalizePhone_(phone);
  var found = ContactsApp.getContactsByPhoneNumber(normalized);
  if (found && found.length > 0) return { ok: true, existed: true };

  var notes = "";
  if (addressLine) notes += "주소: " + addressLine;
  if (mapUrl) notes += (notes ? "\n" : "") + "지도: " + mapUrl;

  var c = ContactsApp.createContact(displayName || normalized, "", notes || "");
  c.addPhone(ContactsApp.Field.MOBILE_PHONE, normalized);

  try {
    if (addressLine) c.addAddress(ContactsApp.Field.HOME_ADDRESS, addressLine);
  } catch (_) {}

  return { ok: true, existed: false };
}

/**
 * ✅ 연락처 동기화(배치)
 * - ✅ 완료/취소 블록 스킵
 * - ✅ 연속 빈 블록 N개면 중단
 * - ✅ 연락처_log에 있으면 ContactsApp 조회 자체 스킵
 */
function syncContactsBatch(isSilent) {
  var sheet = getMainSheet_();
  var blockHeight = getBlockHeight_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.START_ROW) return { summary: "데이터 없음" };

  var stopCtl = makeStopController_();

  var logSheet = getContactLogSheet_();
  var logMap = loadContactLogMap_(logSheet);

  var created = 0, existed = 0, cached = 0, skipped = 0, failed = 0;

  var pendingAppends = [];
  var pendingUpdates = [];

  for (var r = CONFIG.START_ROW; r <= lastRow; r += blockHeight) {
    if (stopCtl.check(sheet, r)) break;

    if (isClosedBlock_(sheet, r)) { skipped++; continue; }

    var nameVal = sheet.getRange(r, CONFIG.POS_NAME.col).getDisplayValue();
    var ignoreNameCheck = !!(CONFIG && CONFIG.CONTACT_IGNORE_NAME_VALIDATION);
    if (!ignoreNameCheck && !isValidName(nameVal)) continue;

    // ✅ 고정 셀에서만 전화번호 가져옴
    var phone = findPhoneInBlock_(sheet, r, blockHeight);

    if (CONFIG.CONTACT_SKIP_IF_NO_PHONE && !phone) {
      skipped++;
      continue;
    }

    var normalized = normalizePhone_(phone);

    // ✅ 주소(F4 + 공백 + F6) / 메모(지도 URL)
    var f4 = sheet.getRange(r + CONFIG.POS_ADDR.row, CONFIG.POS_ADDR.col).getDisplayValue();
    var f6 = sheet.getRange(r + CONFIG.POS_ADDR_EXTRA.row, CONFIG.POS_ADDR_EXTRA.col).getDisplayValue();
    var addressLine = ((f4 || "").toString().trim() + " " + (f6 || "").toString().trim()).replace(/\s+/g, " ").trim();
    var mapUrl = sheet.getRange(r + CONFIG.POS_MAP.row, CONFIG.POS_MAP.col).getDisplayValue();

    var noVal = sheet.getRange(r, CONFIG.POS_NO.col).getDisplayValue();
    var projectLabel = (noVal ? (noVal + " ") : "") + nameVal;

    // ✅ 로그에 있으면 스킵
    if (normalized && logMap[normalized]) {
      cached++;
      var rowNum = logMap[normalized];
      pendingUpdates.push({
        row: rowNum,
        values: [
          normalized, nameVal, projectLabel, addressLine, mapUrl,
          "cached_skip", "", r, new Date()
        ]
      });
      continue;
    }

    try {
      var res = ensureContact_(nameVal, normalized, addressLine, mapUrl);
      if (res.skipped) { skipped++; continue; }

      if (res.existed) existed++; else created++;

      var logRow = [
        normalized, nameVal, projectLabel, addressLine, mapUrl,
        "ok", res.existed ? "Y" : "N", r, new Date()
      ];

      logMap[normalized] = -1;
      pendingAppends.push(logRow);

    } catch (e) {
      failed++;
      if (normalized) {
        pendingAppends.push([
          normalized, nameVal, projectLabel, addressLine, mapUrl,
          "fail: " + (e && e.message ? e.message : e), "", r, new Date()
        ]);
      }
    }
  }

  // 기존행 업데이트
  for (var i = 0; i < pendingUpdates.length; i++) {
    try {
      var u = pendingUpdates[i];
      logSheet.getRange(u.row, 1, 1, 9).setValues([u.values]);
    } catch (_) {}
  }

  // append 배치
  if (pendingAppends.length > 0) {
    var start = logSheet.getLastRow() + 1;
    logSheet.getRange(start, 1, pendingAppends.length, 9).setValues(pendingAppends);
  }

  var summary =
    "생성 " + created +
    " / 기존 " + existed +
    " / 로그스킵 " + cached +
    " / 건너뜀 " + skipped +
    " / 실패 " + failed;

  if (!isSilent) SpreadsheetApp.getUi().alert("✅ 연락처 동기화 완료\n" + summary);
  return { summary: summary };
}
