/** InitSettings.gs
 * clickup settings 시트 자동 생성/보강 + 로더.
 * - 기존 값은 절대 덮어쓰지 않는다.
 * - 누락된 헤더/키/행만 추가한다.
 */

function ensureClickUpSettingsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CLICKUP_SETTINGS.SHEET_NAME);
  if (!sh) sh = ss.insertSheet(CLICKUP_SETTINGS.SHEET_NAME);

  var layout = getClickUpSettingsLayout_();

  ensureSectionTable_(sh, layout.basic, CLICKUP_SETTINGS.BASIC_CONFIG_DEFAULTS);
  ensureSectionTable_(sh, layout.auth, CLICKUP_SETTINGS.CLICKUP_AUTH_DEFAULTS);
  ensureSectionTable_(sh, layout.target, CLICKUP_SETTINGS.CLICKUP_TARGET_DEFAULTS);
  ensureItemMapSection_(sh, layout.itemMap);

  return sh;
}

function getClickUpSettings_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('CLICKUP_SETTINGS_CACHE_V1');
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }

  var sh = ensureClickUpSettingsSheet_();
  var layout = getClickUpSettingsLayout_();

  var basic = readKeyValueSection_(sh, layout.basic);
  var auth = readKeyValueSection_(sh, layout.auth);
  var target = readKeyValueSection_(sh, layout.target);
  var itemMap = readItemMapSection_(sh, layout.itemMap);

  var cfg = {
    sheetName: sh.getName(),
    basic: basic,
    auth: auth,
    target: target,
    itemMap: itemMap
  };

  // LIST_ID는 URL 입력을 허용하고 내부적으로 정규화해서 사용
  cfg.target.CLICKUP_LIST_ID = normalizeClickUpListId_(cfg.target.CLICKUP_LIST_ID);

  cache.put('CLICKUP_SETTINGS_CACHE_V1', JSON.stringify(cfg), 120);
  return cfg;
}

function clearClickUpSettingsCache_() {
  CacheService.getScriptCache().remove('CLICKUP_SETTINGS_CACHE_V1');
}

function validateClickUpRequiredSettings_(settings) {
  var missing = [];
  if (!settings || !settings.auth || !stringValue_(settings.auth.CLICKUP_API_TOKEN)) missing.push('CLICKUP_API_TOKEN');
  if (!settings || !settings.target || !stringValue_(settings.target.CLICKUP_LIST_ID)) {
    missing.push('CLICKUP_LIST_ID');
  } else if (!/^\d+$/.test(stringValue_(settings.target.CLICKUP_LIST_ID))) {
    missing.push('CLICKUP_LIST_ID(숫자 ID 또는 List URL 입력)');
  }
  return missing;
}

function getClickUpSettingsLayout_() {
  return {
    basic: { titleRow: 1, headerRow: 2, dataStartRow: 3, maxRows: 20 },
    auth: { titleRow: 24, headerRow: 25, dataStartRow: 26, maxRows: 10 },
    target: { titleRow: 38, headerRow: 39, dataStartRow: 40, maxRows: 16 },
    itemMap: { titleRow: 60, headerRow: 61, dataStartRow: 62, maxRows: 80 }
  };
}

function ensureSectionTable_(sheet, layout, defaults) {
  var titleCell = sheet.getRange(layout.titleRow, 1);
  if (!titleCell.getDisplayValue()) {
    var titleText = layout.titleRow === 1 ? 'BASIC_CONFIG' : (layout.titleRow === 24 ? 'CLICKUP_AUTH' : 'CLICKUP_TARGET');
    titleCell.setValue(titleText).setFontWeight('bold');
  }

  var headers = ['key', 'value', 'description'];
  var headerRange = sheet.getRange(layout.headerRow, 1, 1, headers.length);
  var existingHeaders = headerRange.getDisplayValues()[0];
  var needHeader = existingHeaders.join('').trim() === '';
  if (needHeader) headerRange.setValues([headers]).setFontWeight('bold');

  var existing = readKeyValueRowsRaw_(sheet, layout);
  var existingMap = {};
  existing.forEach(function(row) {
    var key = stringValue_(row[0]);
    if (key) existingMap[key] = row;
  });

  var appendRows = [];
  defaults.forEach(function(row) {
    var key = row[0];
    if (!existingMap[key]) appendRows.push(row);
  });

  if (appendRows.length > 0) {
    var writeStart = findNextEmptyRow_(sheet, layout.dataStartRow, layout.maxRows, 1);
    sheet.getRange(writeStart, 1, appendRows.length, 3).setValues(appendRows);
  }
}

function ensureItemMapSection_(sheet, layout) {
  var titleCell = sheet.getRange(layout.titleRow, 1);
  if (!titleCell.getDisplayValue()) titleCell.setValue('SYNC_ITEM_MAP').setFontWeight('bold');

  var headerRange = sheet.getRange(layout.headerRow, 1, 1, CLICKUP_SETTINGS.ITEM_MAP_HEADERS.length);
  var existingHeaders = headerRange.getDisplayValues()[0];
  if (existingHeaders.join('').trim() === '') {
    headerRange.setValues([CLICKUP_SETTINGS.ITEM_MAP_HEADERS]).setFontWeight('bold');
  } else {
    // 누락 헤더만 보강
    var merged = [];
    for (var i = 0; i < CLICKUP_SETTINGS.ITEM_MAP_HEADERS.length; i++) {
      merged.push(existingHeaders[i] || CLICKUP_SETTINGS.ITEM_MAP_HEADERS[i]);
    }
    headerRange.setValues([merged]).setFontWeight('bold');
  }

  var currentRows = sheet.getRange(layout.dataStartRow, 1, layout.maxRows, CLICKUP_SETTINGS.ITEM_MAP_HEADERS.length).getDisplayValues();
  var existingCodes = {};
  currentRows.forEach(function(r) {
    var code = stringValue_(r[1]);
    if (code) existingCodes[code] = true;
  });

  var appendRows = [];
  CLICKUP_SETTINGS.ITEM_MAP_DEFAULT_ROWS.forEach(function(row) {
    if (!existingCodes[row[1]]) appendRows.push(row);
  });

  if (appendRows.length > 0) {
    var writeStart = findNextEmptyRow_(sheet, layout.dataStartRow, layout.maxRows, 2);
    sheet.getRange(writeStart, 1, appendRows.length, CLICKUP_SETTINGS.ITEM_MAP_HEADERS.length).setValues(appendRows);
  }
}

function readKeyValueSection_(sheet, layout) {
  var rows = readKeyValueRowsRaw_(sheet, layout);
  var out = {};
  rows.forEach(function(r) {
    var key = stringValue_(r[0]);
    if (!key) return;
    out[key] = r[1];
  });
  return out;
}

function readKeyValueRowsRaw_(sheet, layout) {
  return sheet.getRange(layout.dataStartRow, 1, layout.maxRows, 3).getValues();
}

function readItemMapSection_(sheet, layout) {
  var values = sheet.getRange(layout.dataStartRow, 1, layout.maxRows, CLICKUP_SETTINGS.ITEM_MAP_HEADERS.length).getValues();
  var rows = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var itemCode = stringValue_(row[1]);
    if (!itemCode) continue;

    rows.push({
      useYn: stringValue_(row[0]) || 'Y',
      itemCode: itemCode,
      groupCode: stringValue_(row[2]),
      sourceLabel: stringValue_(row[3]),
      clickupLabel: stringValue_(row[4]) || stringValue_(row[3]),
      dueCol: stringValue_(row[5]),
      dueRowOffset: toIntOrNull_(row[6]),
      doneCol: stringValue_(row[7]),
      doneRowOffset: toIntOrNull_(row[8]),
      noteCol: stringValue_(row[9]),
      noteRowOffset: toIntOrNull_(row[10]),
      linkCol: stringValue_(row[11]),
      linkRowOffset: toIntOrNull_(row[12]),
      sortOrder: Number(row[13]) || 999,
      description: stringValue_(row[14])
    });
  }

  rows.sort(function(a, b) { return a.sortOrder - b.sortOrder; });
  return rows;
}

function findNextEmptyRow_(sheet, startRow, maxRows, keyCol) {
  var values = sheet.getRange(startRow, keyCol, maxRows, 1).getDisplayValues();
  for (var i = 0; i < values.length; i++) {
    if (!stringValue_(values[i][0])) return startRow + i;
  }
  return startRow + maxRows;
}
