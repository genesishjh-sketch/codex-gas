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

  var rows = logSheet.getRange(2, 1, last - 1, 6).getValues();
  for (var i = 0; i < rows.length; i++) {
    var p = (rows[i][0] || "").toString().trim();
    if (!p) continue;
    map[p] = {
      row: i + 2,
      result: (rows[i][5] || "").toString().trim()
    };
  }
  return map;
}

/**
 * 전화번호는 고정셀(D열, 블록시작+2)에서 읽는다
 * - D6, D15, D24 ... 패턴
 */
function findPhoneInBlock_(sheet, blockStartRow, blockHeight) {
  var phoneCellRowOffset = 2; // blockStartRow + 2 => D6 패턴
  var phoneCol = 4;           // D열

  var raw = sheet.getRange(blockStartRow + phoneCellRowOffset, phoneCol).getDisplayValue();
  var s = (raw || "").toString();

  var phoneRegex = /\+?\s*\(?0?1[016789]\)?[\s\-.]?\d{3,4}[\s\-.]?\d{4}/;
  var m = s.match(phoneRegex);
  if (m && m[0]) return normalizePhone_(m[0]);

  // 과거 동작 호환: 고정 셀에 없으면 블록 전체에서 탐색
  var maxCols = Math.min(sheet.getLastColumn(), 40);
  var vals = sheet.getRange(blockStartRow, 1, blockHeight, maxCols).getDisplayValues();
  for (var i = 0; i < vals.length; i++) {
    for (var j = 0; j < vals[i].length; j++) {
      var v = (vals[i][j] || "").toString();
      var m2 = v.match(phoneRegex);
      if (m2 && m2[0]) return normalizePhone_(m2[0]);
    }
  }

  return "";
}

function phoneDigits_(raw) {
  if (!raw) return "";
  var d = raw.toString().replace(/[^\d]/g, "");
  if (d.indexOf("82") === 0 && d.length >= 12) {
    var rest = d.substring(2);
    if (rest.indexOf("10") === 0 && rest.length === 10) d = "0" + rest;
  }
  return d;
}

function isBlockedContactPhone_(raw) {
  var d = phoneDigits_(raw);
  return d === "01000000000";
}

function personHasPhoneDigits_(person, targetDigits) {
  if (!person || !person.phoneNumbers || !targetDigits) return false;
  var td = phoneDigits_(targetDigits);
  for (var i = 0; i < person.phoneNumbers.length; i++) {
    var ph = person.phoneNumbers[i] || {};
    var v1 = phoneDigits_(ph.value || "");
    var v2 = phoneDigits_(ph.canonicalForm || "");
    if (v1 && v1 === td) return true;
    if (v2 && v2 === td) return true;
  }
  return false;
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

var PEOPLE_READ_MASK = "names,phoneNumbers,addresses,biographies,metadata";

var CONTACTS_SERVICE_STATE_ = null;
var PEOPLE_READ_QUOTA_EXCEEDED_ = false;

function isContactsDeprecatedError_(error) {
  var msg = String(error && (error.message || error) || "");
  return msg.indexOf("Contacts API has been deprecated") >= 0;
}

function isPeopleScopeError_(error) {
  var msg = String(error && (error.message || error) || "");
  return msg.indexOf("insufficient authentication scopes") >= 0;
}

function isPeopleApiDisabledError_(error) {
  var msg = String(error && (error.message || error) || "");
  return msg.indexOf("API has not been used") >= 0 ||
    msg.indexOf("is not enabled") >= 0 ||
    msg.indexOf("has not been used in project") >= 0;
}

function isPeopleReadQuotaExceededError_(error) {
  var msg = String(error && (error.message || error) || "");
  return msg.indexOf("Quota exceeded") >= 0 &&
    msg.indexOf("Critical read requests") >= 0;
}

function getPeopleServiceState_() {
  if (typeof People === "undefined" || !People.People) {
    return { ok: false, reason: "people_unavailable" };
  }
  try {
    People.People.get("people/me", { personFields: "names" });
    return { ok: true };
  } catch (e) {
    if (isPeopleScopeError_(e)) {
      return { ok: false, reason: "people_scope" };
    }
    if (isPeopleApiDisabledError_(e)) {
      return { ok: false, reason: "people_disabled" };
    }
    return { ok: false, reason: "people_error" };
  }
}

function getContactsServiceState_() {
  if (CONTACTS_SERVICE_STATE_) return CONTACTS_SERVICE_STATE_;

  // People 객체가 보이면 우선 People 경로를 사용한다.
  // (사전 진단 호출이 실패하더라도 실제 search/create가 동작하는 환경이 있어 선차단하지 않음)
  if (typeof People !== "undefined" && People.People) {
    CONTACTS_SERVICE_STATE_ = { ok: true, provider: "people" };
    return CONTACTS_SERVICE_STATE_;
  }

  var peopleState = getPeopleServiceState_();
  if (typeof ContactsApp === "undefined") {
    CONTACTS_SERVICE_STATE_ = { ok: false, reason: peopleState.reason || "contacts_unavailable" };
    return CONTACTS_SERVICE_STATE_;
  }
  try {
    ContactsApp.getContacts();
    CONTACTS_SERVICE_STATE_ = { ok: true, provider: "contacts" };
    return CONTACTS_SERVICE_STATE_;
  } catch (e) {
    if (isContactsDeprecatedError_(e)) {
      CONTACTS_SERVICE_STATE_ = { ok: false, reason: peopleState.reason || "contacts_deprecated" };
      return CONTACTS_SERVICE_STATE_;
    }
    throw e;
  }
}

function getContactsUnavailableMessage_(reason, actionLabel) {
  var label = actionLabel || "연락처 작업";
  if (reason === "contacts_deprecated") {
    var peopleState = getPeopleServiceState_();
    if (!peopleState.ok && peopleState.reason) {
      reason = peopleState.reason;
    }
  }
  if (reason === "contacts_deprecated") {
    return "⚠️ Contacts API가 종료되어 " + label + " 작업을 건너뜁니다. People API로 이전이 필요합니다.";
  }
  if (reason === "people_scope") {
    return "⚠️ People API 권한이 없어 " + label + " 작업을 건너뜁니다. 고급 서비스 및 스코프를 확인하세요.";
  }
  if (reason === "people_disabled") {
    return "⚠️ People API가 프로젝트에서 비활성화되어 " + label + " 작업을 건너뜁니다. 고급 서비스/Google Cloud 콘솔에서 People API를 활성화하세요.";
  }
  if (reason === "people_unavailable" || reason === "people_error") {
    return "⚠️ 연락처 서비스 상태 확인이 필요하여 " + label + " 작업을 잠시 보류합니다.";
  }
  return "⚠️ ContactsApp을 사용할 수 없어 " + label + " 작업을 건너뜁니다.";
}

function getPeopleServiceMessage_(reason) {
  if (!reason) return "⚠️ People API 상태를 확인할 수 없습니다.";
  if (reason === "people_scope") {
    return "⚠️ People API 권한이 없습니다. 고급 서비스 및 스코프를 확인하세요.";
  }
  if (reason === "people_disabled") {
    return "⚠️ People API가 프로젝트에서 비활성화되어 있습니다. Google Cloud 콘솔에서 People API를 활성화하세요.";
  }
  if (reason === "people_unavailable") {
    return "⚠️ People API 고급 서비스를 사용할 수 없습니다. 서비스가 ON인지 확인하세요.";
  }
  if (reason === "people_error") {
    return "⚠️ People API 호출에 실패했습니다. 권한 재승인 후 다시 시도하세요.";
  }
  return "⚠️ People API 상태: " + reason;
}

function getContactsDiagnosticsSummary_() {
  var messages = [];
  var peopleState = getPeopleServiceState_();
  if (peopleState.ok) {
    messages.push("✅ People API 호출 가능");
  } else {
    messages.push(getPeopleServiceMessage_(peopleState.reason));
  }

  var contactsState = getContactsServiceState_();
  if (contactsState.ok) {
    messages.push("✅ 연락처 서비스: " + (contactsState.provider === "people" ? "People API" : "ContactsApp"));
  } else {
    messages.push(getContactsUnavailableMessage_(contactsState.reason, "연락처 작업"));
  }

  if (typeof ContactsApp === "undefined") {
    messages.push("ℹ️ ContactsApp 객체가 없습니다. (Apps Script 기본 제공 서비스)");
  }

  return messages.join("\n");
}

function assertPeopleEnabled_() {
  if (typeof People === "undefined" || !People.People) {
    throw new Error("People API를 사용할 수 없습니다. Apps Script 고급 서비스에서 People API를 활성화하세요.");
  }
}

function findPeopleContactsByPhone_(normalizedPhone) {
  var person = findContactPersonByPhone_(normalizedPhone);
  return person ? [person] : [];
}

function findContactPersonByPhone_(phone) {
  assertPeopleEnabled_();
  var digits = phoneDigits_(phone);
  if (!digits) return null;

  if (PEOPLE_READ_QUOTA_EXCEEDED_) return null;

  function search_(query) {
    if (!query) return [];
    var response;
    try {
      response = People.People.searchContacts({
        query: query,
        readMask: PEOPLE_READ_MASK,
        pageSize: 30
      });
    } catch (e) {
      if (isPeopleReadQuotaExceededError_(e)) {
        PEOPLE_READ_QUOTA_EXCEEDED_ = true;
        return [];
      }
      try {
        response = People.People.searchContacts({
          query: query,
          personFields: PEOPLE_READ_MASK,
          pageSize: 30
        });
      } catch (e2) {
        if (isPeopleReadQuotaExceededError_(e2)) {
          PEOPLE_READ_QUOTA_EXCEEDED_ = true;
          return [];
        }
        throw e2;
      }
    }
    var results = (response && response.results) ? response.results : [];
    var matched = [];
    for (var i = 0; i < results.length; i++) {
      var person = results[i].person || results[i];
      if (personHasPhoneDigits_(person, digits)) matched.push(person);
    }
    return matched;
  }

  var byDigits = search_(digits);
  if (byDigits.length > 0) return byDigits[0];

  var normalized = normalizePhone_(digits);
  var byNormalized = search_(normalized);
  return byNormalized.length > 0 ? byNormalized[0] : null;
}

function findContactsByPhone_(normalizedPhone) {
  if (!normalizedPhone) return [];
  var state = getContactsServiceState_();
  if (!state.ok) return [];
  if (state.provider === "people" && PEOPLE_READ_QUOTA_EXCEEDED_) return [];
  if (state.provider === "people") {
    return findPeopleContactsByPhone_(normalizedPhone);
  }
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
  var digits = phoneDigits_(normalized);
  if (!digits) return { ok: false, skipped: true, reason: "invalid_phone" };
  if (isBlockedContactPhone_(digits)) {
    return { ok: false, skipped: true, reason: "blocked_phone" };
  }

  var safeName = (displayName || normalized).toString().trim() || normalized;
  var safeAddress = (addressLine || "").toString().trim();
  var safeMapUrl = (mapUrl || "").toString().trim();

  try {
    if (state.provider === "people") {
      var found = findContactPersonByPhone_(digits);
      var updateFields = ["names", "phoneNumbers"];
      if (safeAddress) updateFields.push("addresses");
      if (safeMapUrl) updateFields.push("biographies");

      if (found && found.resourceName) {
        var full = People.People.get(found.resourceName, { personFields: PEOPLE_READ_MASK });
        var patch = {
          resourceName: full.resourceName,
          etag: full.etag,
          metadata: full.metadata,
          names: [{ givenName: safeName }],
          phoneNumbers: [{ value: normalized, type: "mobile" }]
        };
        if (safeAddress) patch.addresses = [{ formattedValue: safeAddress, type: "work" }];
        if (safeMapUrl) patch.biographies = [{ value: safeMapUrl, contentType: "TEXT_PLAIN" }];

        People.People.updateContact(patch, full.resourceName, {
          updatePersonFields: updateFields.join(",")
        });
        return { ok: true, existed: true, updated: true };
      }

      var person = {
        names: [{ givenName: safeName }],
        phoneNumbers: [{ value: normalized, type: "mobile" }]
      };
      if (safeAddress) person.addresses = [{ formattedValue: safeAddress, type: "work" }];
      if (safeMapUrl) person.biographies = [{ value: safeMapUrl, contentType: "TEXT_PLAIN" }];

      People.People.createContact(person);
      return { ok: true, existed: false, created: true, readQuotaBypass: PEOPLE_READ_QUOTA_EXCEEDED_ };
    }

    var notes = "";
    if (safeAddress) notes += "주소: " + safeAddress;
    if (safeMapUrl) notes += (notes ? "\n" : "") + "지도: " + safeMapUrl;

    var foundContacts = findContactsByPhone_(normalized);
    if (foundContacts && foundContacts.length > 0) return { ok: true, existed: true };

    var c = ContactsApp.createContact(safeName, "", notes || "");
    c.addPhone(ContactsApp.Field.MOBILE_PHONE, normalized);
    try {
      if (safeAddress) c.addAddress(ContactsApp.Field.HOME_ADDRESS, safeAddress);
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
    var unavailableMessage = getContactsUnavailableMessage_(contactsState.reason, "연락처 동기화");
    if (!isSilent) {
      SpreadsheetApp.getUi().alert(unavailableMessage);
    }
    return { summary: unavailableMessage };
  }

  var stopCtl = makeStopController_();

  var logSheet = getContactLogSheet_();
  var logMap = loadContactLogMap_(logSheet);

  var created = 0, existed = 0, cached = 0, skipped = 0, failed = 0;
  var quotaBypass = 0;

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
    var logEntry = normalized ? logMap[normalized] : null;
    var logRowNum = logEntry ? logEntry.row : null;
    var logResult = logEntry ? logEntry.result : "";
    var shouldSkipByLog = !!(logRowNum && CONFIG.CONTACT_SKIP_IF_LOGGED !== false &&
      (logResult === "ok" || logResult === "cached_skip"));
    var shouldVerifyLogged = CONFIG.CONTACT_VERIFY_LOGGED !== false;
    var canSkip = shouldSkipByLog;
    if (shouldSkipByLog && shouldVerifyLogged && normalized && !PEOPLE_READ_QUOTA_EXCEEDED_) {
      var existing = findContactsByPhone_(normalized);
      if (!existing || existing.length === 0) {
        canSkip = false;
      }
    }
    if (canSkip) {
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

      if (res.readQuotaBypass) quotaBypass++;

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

  if (quotaBypass > 0) {
    summary += " / 읽기쿼터우회 " + quotaBypass;
  }

  if (PEOPLE_READ_QUOTA_EXCEEDED_) {
    summary += "\n⚠️ People API 읽기 쿼터 초과로 검색/검증을 건너뛰고 신규 등록 위주로 처리했습니다.";
  }

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
    var unavailableMessage = getContactsUnavailableMessage_(state.reason, "연락처 점검");
    if (!isSilent) {
      SpreadsheetApp.getUi().alert(unavailableMessage);
    }
    return { summary: unavailableMessage };
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
