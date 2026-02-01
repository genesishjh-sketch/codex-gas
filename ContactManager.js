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

function getContactBlockInfo_(sheet, blockStartRow, blockHeight, nameVal) {
  var nameValue = (typeof nameVal !== "undefined") ? nameVal : sheet.getRange(blockStartRow, CONFIG.POS_NAME.col).getDisplayValue();
  var phone = findPhoneInBlock_(sheet, blockStartRow, blockHeight);
  var normalized = normalizePhone_(phone);
  var f4 = sheet.getRange(blockStartRow + CONFIG.POS_ADDR.row, CONFIG.POS_ADDR.col).getDisplayValue();
  var f6 = sheet.getRange(blockStartRow + CONFIG.POS_ADDR_EXTRA.row, CONFIG.POS_ADDR_EXTRA.col).getDisplayValue();
  var addressLine = ((f4 || "").toString().trim() + " " + (f6 || "").toString().trim()).replace(/\s+/g, " ").trim();
  var mapUrl = sheet.getRange(blockStartRow + CONFIG.POS_MAP.row, CONFIG.POS_MAP.col).getDisplayValue();
  var noVal = sheet.getRange(blockStartRow, CONFIG.POS_NO.col).getDisplayValue();
  var projectLabel = (noVal ? (noVal + " ") : "") + nameValue;

  return {
    nameVal: nameValue,
    phone: phone,
    normalized: normalized,
    addressLine: addressLine,
    mapUrl: mapUrl,
    projectLabel: projectLabel
  };
}

function queueContactLogAppend_(pendingAppends, info, result, sourceRow, existedFlag) {
  pendingAppends.push([
    info.normalized || "",
    info.nameVal || "",
    info.projectLabel || "",
    info.addressLine || "",
    info.mapUrl || "",
    result,
    existedFlag || "",
    sourceRow,
    new Date()
  ]);
}

var CONTACTS_SERVICE_STATE_ = null;

function isContactsDeprecatedError_(error) {
  var msg = String(error && (error.message || error) || "");
  return msg.indexOf("Contacts API has been deprecated") >= 0;
}

function getContactsServiceState_() {
  if (CONTACTS_SERVICE_STATE_) return CONTACTS_SERVICE_STATE_;
  if (typeof ContactsApp === "undefined") {
    CONTACTS_SERVICE_STATE_ = { ok: false, reason: "contacts_unavailable" };
    return CONTACTS_SERVICE_STATE_;
  }
  try {
    ContactsApp.getContacts();
    CONTACTS_SERVICE_STATE_ = { ok: true };
    return CONTACTS_SERVICE_STATE_;
  } catch (e) {
    if (isContactsDeprecatedError_(e)) {
      CONTACTS_SERVICE_STATE_ = { ok: false, reason: "contacts_deprecated" };
      return CONTACTS_SERVICE_STATE_;
    }
    throw e;
  }
}

function findContactsByPhone_(normalizedPhone) {
  if (!normalizedPhone) return [];
  var state = getContactsServiceState_();
  if (!state.ok) return [];
  if (ContactsApp && typeof ContactsApp.getContactsByPhoneNumber === "function") {
    return ContactsApp.getContactsByPhoneNumber(normalizedPhone) || [];
  }

  var contacts = ContactsApp.getContacts();
  var matches = [];
  for (var i = 0; i < contacts.length; i++) {
    var phones = contacts[i].getPhones();
    for (var j = 0; j < phones.length; j++) {
      var value = phones[j].getPhoneNumber();
      if (normalizePhone_(value) === normalizedPhone) {
        matches.push(contacts[i]);
        break;
      }
    }
  }
  return matches;
}

function ensureContact_(displayName, phone, addressLine, mapUrl) {
  if (!phone) return { ok: false, skipped: true, reason: "no_phone" };
  var state = getContactsServiceState_();
  if (!state.ok) return { ok: false, skipped: true, reason: state.reason };

  var normalized = normalizePhone_(phone);
  var found = findContactsByPhone_(normalized);
  if (found && found.length > 0) return { ok: true, existed: true };

  var notes = "";
  if (addressLine) notes += "주소: " + addressLine;
  if (mapUrl) notes += (notes ? "\n" : "") + "지도: " + mapUrl;

  try {
    var c = ContactsApp.createContact(displayName || normalized, "", notes || "");
    c.addPhone(ContactsApp.Field.MOBILE_PHONE, normalized);

    try {
      if (addressLine) c.addAddress(ContactsApp.Field.HOME_ADDRESS, addressLine);
    } catch (_) {}
  } catch (e) {
    if (isContactsDeprecatedError_(e)) {
      return { ok: false, skipped: true, reason: "contacts_deprecated" };
    }
    throw e;
  }

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

  var contactsState = getContactsServiceState_();
  if (!contactsState.ok) {
    if (!isSilent) {
      var msg = (contactsState.reason === "contacts_deprecated") ?
        "⚠️ Contacts API가 종료되어 연락처 동기화를 건너뜁니다. People API로 이전이 필요합니다." :
        "⚠️ ContactsApp을 사용할 수 없어 연락처 동기화를 건너뜁니다.";
      SpreadsheetApp.getUi().alert(msg);
    }
    return { summary: contactsState.reason || "ContactsApp unavailable" };
  }

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
    if (!ignoreNameCheck && !isValidName(nameVal)) {
      if (CONFIG.CONTACT_LOG_SKIP_REASONS) {
        var invalidInfo = getContactBlockInfo_(sheet, r, blockHeight, nameVal);
        queueContactLogAppend_(pendingAppends, invalidInfo, "skip: invalid_name", r, "");
      }
      continue;
    }

    var info = getContactBlockInfo_(sheet, r, blockHeight, nameVal);
    var phone = info.phone;

    if (CONFIG.CONTACT_SKIP_IF_NO_PHONE && !phone) {
      if (CONFIG.CONTACT_LOG_SKIP_REASONS) {
        queueContactLogAppend_(pendingAppends, info, "skip: no_phone", r, "");
      }
      skipped++;
      continue;
    }

    var normalized = info.normalized;

    // ✅ 로그에 있으면 스킵
    var logRowNum = normalized ? logMap[normalized] : null;
    if (logRowNum && CONFIG.CONTACT_SKIP_IF_LOGGED !== false) {
      cached++;
      pendingUpdates.push({
        row: logRowNum,
        values: [
          normalized, info.nameVal, info.projectLabel, info.addressLine, info.mapUrl,
          "cached_skip", "", r, new Date()
        ]
      });
      continue;
    }

    try {
      var res = ensureContact_(info.nameVal, normalized, info.addressLine, info.mapUrl);
      if (res.skipped) {
        if (CONFIG.CONTACT_LOG_SKIP_REASONS && res.reason) {
          queueContactLogAppend_(pendingAppends, info, "skip: " + res.reason, r, "");
        }
        skipped++;
        continue;
      }

      if (res.existed) existed++; else created++;

      if (logRowNum) {
        pendingUpdates.push({
          row: logRowNum,
          values: [
            normalized, info.nameVal, info.projectLabel, info.addressLine, info.mapUrl,
            "ok", res.existed ? "Y" : "N", r, new Date()
          ]
        });
      } else {
        logMap[normalized] = -1;
        queueContactLogAppend_(pendingAppends, info, "ok", r, res.existed ? "Y" : "N");
      }

    } catch (e) {
      failed++;
      if (normalized) {
        queueContactLogAppend_(pendingAppends, info, "fail: " + (e && e.message ? e.message : e), r, "");
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

/**
 * ✅ 연락처_log 기준 실재 연락처 존재 여부 검사
 * - 연락처_log에 있으나 ContactsApp에 없으면 RESULT를 missing_in_contacts로 표기
 */
function auditContactLog_(isSilent) {
  var logSheet = getContactLogSheet_();
  var state = getContactsServiceState_();
  if (!state.ok) {
    if (!isSilent) {
      var msg = (state.reason === "contacts_deprecated") ?
        "⚠️ Contacts API가 종료되어 점검을 건너뜁니다. People API로 이전이 필요합니다." :
        "⚠️ ContactsApp을 사용할 수 없어 점검을 건너뜁니다.";
      SpreadsheetApp.getUi().alert(msg);
    }
    return { summary: state.reason || "ContactsApp unavailable" };
  }
  var lastRow = logSheet.getLastRow();
  if (lastRow < 2) {
    if (!isSilent) SpreadsheetApp.getUi().alert("ℹ️ 연락처_log 데이터가 없습니다.");
    return { summary: "로그 데이터 없음" };
  }

  var data = logSheet.getRange(2, 1, lastRow - 1, 9).getValues();
  var updates = [];
  var checked = 0;
  var missing = 0;
  var present = 0;
  var skipped = 0;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var phone = (row[0] || "").toString().trim();
    if (!phone) {
      skipped++;
      continue;
    }
    var normalized = normalizePhone_(phone);
    checked++;

    var found = findContactsByPhone_(normalized);
    if (found && found.length > 0) {
      present++;
      continue;
    }

    missing++;
    updates.push({
      row: i + 2,
      values: [
        normalized,
        row[1] || "",
        row[2] || "",
        row[3] || "",
        row[4] || "",
        "missing_in_contacts",
        row[6] || "",
        row[7] || "",
        new Date()
      ]
    });
  }

  for (var u = 0; u < updates.length; u++) {
    var item = updates[u];
    try {
      logSheet.getRange(item.row, 1, 1, 9).setValues([item.values]);
    } catch (_) {}
  }

  var summary =
    "검사 " + checked +
    " / 존재 " + present +
    " / 누락 " + missing +
    " / 스킵 " + skipped;

  if (!isSilent) SpreadsheetApp.getUi().alert("✅ 연락처 로그 점검 완료\n" + summary);
  return { summary: summary };
}
