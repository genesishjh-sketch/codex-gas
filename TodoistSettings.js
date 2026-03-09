/** settings 시트에서 Todoist 동기화 설정/매핑을 읽는 모듈 */
function getTodoistSyncSettings_() {
  var sheet = getTodoistSettingsSheet_();
  var values = sheet.getDataRange().getDisplayValues();

  var defaults = {
    todoist_project_id: '',
    sync_target_sheet: TODOIST_SYNC.DEFAULT_TARGET_SHEET,
    due_date_field: 'plan_date',
    task_title_template: 'project_name&" | "&step_name&" 예정"',
    label_template: '',
    exclude_done: true,
    realtime_sync: true,
    use_assignee: true,
    use_description: false,
    use_labels: false
  };

  var settings = {};
  Object.keys(defaults).forEach(function(key) { settings[key] = defaults[key]; });

  for (var i = 0; i < values.length; i++) {
    var key = (values[i][0] || '').toString().trim();
    var val = values[i][1];
    if (!key || !defaults.hasOwnProperty(key)) continue;

    if (typeof defaults[key] === 'boolean') {
      settings[key] = parseBoolean_(val, defaults[key]);
    } else {
      settings[key] = (val || '').toString().trim();
    }
  }

  return settings;
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
  if (!mapped.todoist_user_id) return null;
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
