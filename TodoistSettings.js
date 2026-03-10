/** settings 시트에서 Todoist 동기화 설정/매핑을 읽는 모듈 */
function getTodoistSyncSettings_() {
  var sheet = ensureTodoistSettingsLayout_();
  var values = sheet.getDataRange().getDisplayValues();

  var defaults = getTodoistSettingDefaults_();

  var settings = {};
  Object.keys(defaults).forEach(function(key) { settings[key] = defaults[key]; });

  var normalizedDefaults = {};
  Object.keys(defaults).forEach(function(key) {
    normalizedDefaults[normalizeSettingKey_(key)] = key;
  });

  for (var i = 0; i < values.length; i++) {
    var key = (values[i][0] || '').toString().trim();
    var val = values[i][1];
    if (!key) continue;

    var normalizedKey = normalizeSettingKey_(key);
    var matchedKey = normalizedDefaults[normalizedKey] || getSettingsKeyAlias_(normalizedKey);
    if (!matchedKey || !defaults.hasOwnProperty(matchedKey)) continue;

    if (typeof defaults[matchedKey] === 'boolean') {
      settings[matchedKey] = parseBoolean_(val, defaults[matchedKey]);
    } else {
      settings[matchedKey] = (val || '').toString().trim();
    }
  }

  return settings;
}

function getTodoistSettingDefaults_() {
  var defaults = {};

  TODOIST_SETTINGS_LAYOUT.sections.forEach(function(section) {
    if (section.type !== 'keyValue' || !section.rows) return;

    section.rows.forEach(function(row) {
      defaults[row.key] = row.defaultValue;
    });
  });

  return defaults;
}

function ensureTodoistSettingsLayout_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TODOIST_SYNC.SETTINGS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TODOIST_SYNC.SETTINGS_SHEET_NAME);
  }

  var values = sheet.getDataRange().getDisplayValues();
  var existingKeyValue = {};

  values.forEach(function(row) {
    var key = normalizeSettingKey_(row[0]);
    if (!key) return;
    existingKeyValue[key] = row[1];
  });

  var outputRows = [];

  TODOIST_SETTINGS_LAYOUT.sections.forEach(function(section, sectionIndex) {
    if (sectionIndex > 0) outputRows.push(['', '', '', '', '']);
    outputRows.push([section.title, '', '', '', '']);

    if (section.type === 'keyValue') {
      outputRows.push(TODOIST_SETTINGS_LAYOUT.columns);
      section.rows.forEach(function(row) {
        var normalizedKey = normalizeSettingKey_(row.key);
        var existingValue = existingKeyValue[normalizedKey];
        var value = existingValue !== undefined && existingValue !== '' ? existingValue : row.defaultValue;
        outputRows.push([row.key, value, row.description || '', row.example || '', '']);
      });
      return;
    }

    if (section.type === 'table') {
      var header = section.header || [];
      outputRows.push(header);

      var existingRows = extractMappingRowsFromValues_(values, header);
      var rowsToWrite = existingRows.length > 0 ? existingRows : (section.rows || []);
      rowsToWrite.forEach(function(row) {
        outputRows.push(row);
      });
    }
  });

  var maxWidth = outputRows.reduce(function(max, row) {
    return Math.max(max, row.length);
  }, 1);

  var normalizedRows = outputRows.map(function(row) {
    var copy = row.slice();
    while (copy.length < maxWidth) copy.push('');
    return copy;
  });

  sheet.clear();
  sheet.getRange(1, 1, normalizedRows.length, maxWidth).setValues(normalizedRows);
  sheet.getRange(1, 1, normalizedRows.length, 1).setFontWeight('bold');
  sheet.getRange(1, 1, normalizedRows.length, maxWidth).setVerticalAlignment('middle');
  sheet.autoResizeColumns(1, maxWidth);

  return sheet;
}

function extractMappingRowsFromValues_(values, header) {
  if (!values || !values.length || !header || !header.length) return [];

  var normalizedHeader = header.map(function(v) { return normalizeSettingKey_(v); });
  var headerRowIndex = -1;

  for (var i = 0; i < values.length; i++) {
    var match = true;
    for (var j = 0; j < normalizedHeader.length; j++) {
      if (normalizeSettingKey_(values[i][j]) !== normalizedHeader[j]) {
        match = false;
        break;
      }
    }
    if (match) {
      headerRowIndex = i;
      break;
    }
  }

  if (headerRowIndex < 0) return [];

  var rows = [];
  for (var r = headerRowIndex + 1; r < values.length; r++) {
    var row = values[r];
    var isEmpty = true;
    for (var c = 0; c < normalizedHeader.length; c++) {
      if ((row[c] || '').toString().trim() !== '') {
        isEmpty = false;
        break;
      }
    }
    if (isEmpty) break;
    rows.push(row.slice(0, normalizedHeader.length));
  }

  return rows;
}

function normalizeSettingKey_(value) {
  return (value || '')
    .toString()
    .trim()
    .toLowerCase()
    .replace(/[\u200B-\u200D\uFEFF]/g, '');
}

function getSettingsKeyAlias_(normalizedKey) {
  var aliases = {
    todoist_token: 'todoist_api_token',
    todoistapitoken: 'todoist_api_token',
    todoist_api_key: 'todoist_api_token'
  };
  return aliases[normalizedKey] || '';
}

function getTodoistApiToken_() {
  var settingsToken = (getTodoistSyncSettings_().todoist_api_token || '').toString().trim();
  if (settingsToken) {
    return {
      token: settingsToken,
      source: 'settings'
    };
  }

  var scriptProperties = PropertiesService.getScriptProperties();
  var scriptToken = (scriptProperties.getProperty(TODOIST_SYNC.PROPERTY_API_TOKEN) || '').toString().trim();
  if (!scriptToken) {
    scriptToken = (scriptProperties.getProperty('TODOIST_TOKEN') || '').toString().trim();
  }
  if (scriptToken) {
    return {
      token: scriptToken,
      source: 'script_properties'
    };
  }

  return {
    token: '',
    source: ''
  };
}

function getSectionMappingMap_() {
  var table = readMappingBlockByHeader_(getTodoistSettingsSheet_(), ['section값', 'todoist_section_id']);
  var map = {};

  table.rows.forEach(function(row) {
    var sectionName = (row[0] || '').toString().trim();
    var todoistSectionId = (row[1] || '').toString().trim();
    if (!sectionName || !todoistSectionId) return;
    map[sectionName] = todoistSectionId;
  });

  return map;
}

function getManagerMappingMap_() {
  var table = readMappingBlockByHeader_(getTodoistSettingsSheet_(), ['manager_name', 'todoist_user_email', 'todoist_user_id', 'active']);
  var map = {};

  table.rows.forEach(function(row) {
    var managerName = (row[0] || '').toString().trim();
    if (!managerName) return;

    map[managerName] = {
      manager_name: managerName,
      todoist_user_email: (row[1] || '').toString().trim(),
      todoist_user_id: (row[2] || '').toString().trim(),
      active: parseBoolean_(row[3], true)
    };
  });

  return map;
}

function getStepProjectMappingRules_() {
  var table = readMappingBlockByHeader_(getTodoistSettingsSheet_(), ['match_type', 'pattern', 'todoist_project_id']);
  var rules = [];

  table.rows.forEach(function(row, index) {
    var matchType = normalizeStepProjectMatchType_((row[0] || '').toString().trim());
    var pattern = (row[1] || '').toString().trim();
    var projectId = (row[2] || '').toString().trim();
    var priorityRaw = parseInt((row[3] || '').toString().trim(), 10);
    var priority = isNaN(priorityRaw) ? (1000 + index) : priorityRaw;
    var active = parseBoolean_(row[4], true);

    if (!active || !matchType || !pattern || !projectId) return;

    rules.push({
      match_type: matchType,
      pattern: pattern,
      todoist_project_id: projectId,
      priority: priority
    });
  });

  rules.sort(function(a, b) {
    if (a.priority === b.priority) return 0;
    return a.priority < b.priority ? -1 : 1;
  });

  return rules;
}

function normalizeStepProjectMatchType_(matchType) {
  var normalized = (matchType || '').toString().trim().toLowerCase();
  if (normalized === 'exact' || normalized === 'contains' || normalized === 'regex') return normalized;
  return '';
}

function getTodoistSectionIdBySection_(sectionValue, sectionMap) {
  var key = (sectionValue || '').toString().trim();
  if (!key) return '';
  return sectionMap[key] || '';
}

function getTodoistAssigneeByManager_(managerName, managerMap) {
  var key = (managerName || '').toString().trim();
  if (!key) return null;

  var mapped = managerMap[key];
  if (!mapped || !mapped.active) return null;
  return mapped;
}

function getTodoistSettingsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TODOIST_SYNC.SETTINGS_SHEET_NAME);
  if (!sheet) throw new Error('settings 시트를 찾을 수 없습니다: ' + TODOIST_SYNC.SETTINGS_SHEET_NAME);
  return sheet;
}

function readMappingBlockByHeader_(sheet, headerCandidates) {
  var data = sheet.getDataRange().getDisplayValues();
  var normalizedCandidates = headerCandidates.map(function(v) { return v.toString().trim().toLowerCase(); });

  var headerRowIndex = -1;
  for (var i = 0; i < data.length; i++) {
    var current = (data[i][0] || '').toString().trim().toLowerCase();
    if (current === normalizedCandidates[0]) {
      var allMatched = true;
      for (var j = 0; j < normalizedCandidates.length; j++) {
        var cell = (data[i][j] || '').toString().trim().toLowerCase();
        if (cell !== normalizedCandidates[j]) {
          allMatched = false;
          break;
        }
      }
      if (allMatched) {
        headerRowIndex = i;
        break;
      }
    }
  }

  if (headerRowIndex < 0) {
    throw new Error('settings에서 매핑 헤더를 찾지 못했습니다: ' + headerCandidates.join(', '));
  }

  var rows = [];
  for (var r = headerRowIndex + 1; r < data.length; r++) {
    var first = (data[r][0] || '').toString().trim();
    var second = (data[r][1] || '').toString().trim();
    if (!first && !second) break;
    rows.push(data[r]);
  }

  return { headerRowIndex: headerRowIndex + 1, rows: rows };
}

function parseBoolean_(value, defaultValue) {
  if (value === true || value === false) return value;
  var normalized = (value || '').toString().trim().toLowerCase();
  if (normalized === 'true' || normalized === '1' || normalized === 'y' || normalized === 'yes') return true;
  if (normalized === 'false' || normalized === '0' || normalized === 'n' || normalized === 'no') return false;
  return !!defaultValue;
}
