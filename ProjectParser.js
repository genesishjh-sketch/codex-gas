/** ProjectParser.gs
 * 블록 단위 데이터 파싱.
 * - 절대좌표 하드코딩 대신 startRow + rowOffset 구조 사용.
 */

function parseProjectBlock_(sheet, startRow, settings) {
  var basic = settings.basic || {};
  var repOffset = Number(basic.REPRESENTATIVE_ROW_OFFSET || 1);
  var repRow = startRow + repOffset;
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'Etc/UTC';

  function gv(col, rowOffset) {
    return sheet.getRange(startRow + rowOffset, colToNumber_(col)).getDisplayValue();
  }
  function vv(col, rowOffset) {
    return sheet.getRange(startRow + rowOffset, colToNumber_(col)).getValue();
  }

  var title = compressSpace_(gv('C', 0));
  var customerName = compressSpace_(gv('D', 1));
  var addr = joinNonEmpty_([gv('F', 0), gv('F', 2)], ' ');

  var folderPairs = [];
  for (var i = 0; i <= 6; i++) {
    var folderName = compressSpace_(gv('R', i));
    var folderLink = compressSpace_(gv('S', i));
    if (!folderName || !folderLink) continue;
    folderPairs.push({ name: folderName, link: folderLink });
  }

  return {
    sheetName: sheet.getName(),
    blockStartRow: startRow,
    representativeRow: repRow,
    timeZone: tz,
    projectTitle: title,
    customerName: customerName,
    customerPhone: compressSpace_(gv('D', 2)),
    projectType: compressSpace_(gv('D', 3)),
    contractDate: vv('D', 4),
    balanceDate: vv('D', 5),
    listDeadline: vv('D', 6),
    measureDate: vv('H', 1),
    address: addr,
    password: compressSpace_(gv('F', 3)),
    addressLink: compressSpace_(gv('F', 4)),
    extraMemo: compressSpace_(gv('F', 7)),
    homeConsultMemo: compressSpace_(gv('G', 7)),
    constructionConsultMemo: compressSpace_(gv('M', 7)),
    viewer3dLink: compressSpace_(gv('K', 0)),
    edit3dLink: compressSpace_(gv('K', 3)),
    shoppingListLink: compressSpace_(gv('K', 4)),
    folderLinks: folderPairs
  };
}

function buildProjectTitleFallback_(projectData, projectUniqueId, settings) {
  var includeProjectId = String(
    settings && settings.target && settings.target.CLICKUP_PARENT_TITLE_INCLUDE_PROJECT_ID || 'FALSE'
  ).toUpperCase() === 'TRUE';

  var location = parseFirstParenText_(projectData.projectTitle) || parseFirstParenText_(projectData.address) || '현장';
  var baseTitle = projectData.projectTitle || (projectData.customerName + '(' + location + ')');

  var rightParts = [];
  if (projectData.customerName) rightParts.push(projectData.customerName);
  if (projectData.measureDate) rightParts.push('실측 ' + formatDateYmd_(projectData.measureDate, projectData.timeZone));
  if (projectData.customerPhone) rightParts.push(projectData.customerPhone);
  if (projectData.address) rightParts.push(projectData.address);

  var name = baseTitle;
  if (rightParts.length > 0) name += ' | ' + rightParts.join(' | ');
  if (includeProjectId) name = '[' + projectUniqueId + '] ' + name;
  return name;
}

function buildParentTaskDescription_(projectData, projectUniqueId) {
  var lines = [];
  lines.push('# 프로젝트 요약');
  lines.push('- project_unique_id: ' + projectUniqueId);
  lines.push('- 프로젝트 제목: ' + (projectData.projectTitle || ''));
  lines.push('- 고객명: ' + (projectData.customerName || ''));
  lines.push('- 연락처: ' + (projectData.customerPhone || ''));
  lines.push('- 타입: ' + (projectData.projectType || ''));
  lines.push('- 계약일: ' + formatDateYmd_(projectData.contractDate, projectData.timeZone));
  lines.push('- 잔금일: ' + formatDateYmd_(projectData.balanceDate, projectData.timeZone));
  lines.push('- 리스트마감: ' + formatDateYmd_(projectData.listDeadline, projectData.timeZone));
  lines.push('- 주소: ' + (projectData.address || ''));
  lines.push('- 비번: ' + (projectData.password || ''));
  lines.push('- 주소링크: ' + (projectData.addressLink || ''));
  lines.push('- 하단 추가설명: ' + (projectData.extraMemo || ''));
  lines.push('- 홈스타일 상담 메모: ' + (projectData.homeConsultMemo || ''));
  lines.push('- 시공 상담 메모: ' + (projectData.constructionConsultMemo || ''));
  lines.push('- 3D 뷰어 링크: ' + (projectData.viewer3dLink || ''));
  lines.push('- 3D 수정 링크: ' + (projectData.edit3dLink || ''));
  lines.push('- 구매리스트 링크: ' + (projectData.shoppingListLink || ''));
  lines.push('- sheet name: ' + projectData.sheetName);
  lines.push('- block start row: ' + projectData.blockStartRow);
  lines.push('- representative row: ' + projectData.representativeRow);

  lines.push('');
  lines.push('## 폴더/링크');
  if (!projectData.folderLinks || projectData.folderLinks.length === 0) {
    lines.push('- 없음');
  } else {
    projectData.folderLinks.forEach(function(p) {
      lines.push('- ' + p.name + ': ' + p.link);
    });
  }

  lines.push('');
  lines.push('[SYNC_META]');
  lines.push('project_unique_id=' + projectUniqueId);
  lines.push('sheet_name=' + projectData.sheetName);
  lines.push('block_start_row=' + projectData.blockStartRow);
  lines.push('representative_row=' + projectData.representativeRow);
  lines.push('[/SYNC_META]');

  return lines.join('\n');
}

function buildSubtaskPayloads_(sheet, startRow, projectUniqueId, settings) {
  var itemMap = (settings.itemMap || []).filter(function(r) {
    var flag = stringValue_(r.useYn).toUpperCase();
    return flag !== 'N' && flag !== 'FALSE';
  });

  var payloads = [];
  itemMap.forEach(function(r) {
    var dueRow = r.dueRowOffset === null ? null : startRow + r.dueRowOffset;
    var doneRow = r.doneRowOffset === null ? null : startRow + r.doneRowOffset;
    var noteRow = r.noteRowOffset === null ? null : startRow + r.noteRowOffset;
    var linkRow = r.linkRowOffset === null ? null : startRow + r.linkRowOffset;

    var dueCol = colToNumber_(r.dueCol);
    var doneCol = colToNumber_(r.doneCol);
    var noteCol = colToNumber_(r.noteCol);
    var linkCol = colToNumber_(r.linkCol);

    var dueValue = (dueRow && dueCol) ? sheet.getRange(dueRow, dueCol).getValue() : '';
    var doneValue = (doneRow && doneCol) ? sheet.getRange(doneRow, doneCol).getValue() : '';
    var noteValue = (noteRow && noteCol) ? sheet.getRange(noteRow, noteCol).getDisplayValue() : '';
    var linkValue = (linkRow && linkCol) ? sheet.getRange(linkRow, linkCol).getDisplayValue() : '';

    var itemSyncKey = buildItemSyncKey_(projectUniqueId, r.itemCode);
    var description = buildSubtaskDescription_({
      projectUniqueId: projectUniqueId,
      itemSyncKey: itemSyncKey,
      groupCode: r.groupCode,
      sourceLabel: r.sourceLabel,
      dueCell: toRelativeCellA1_(r.dueCol, r.dueRowOffset),
      doneCell: toRelativeCellA1_(r.doneCol, r.doneRowOffset),
      noteCell: toRelativeCellA1_(r.noteCol, r.noteRowOffset),
      linkCell: toRelativeCellA1_(r.linkCol, r.linkRowOffset),
      dueValue: dueValue,
      doneValue: doneValue,
      noteValue: noteValue,
      linkValue: linkValue,
      tz: SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'Etc/UTC'
    });

    payloads.push({
      itemCode: r.itemCode,
      groupCode: r.groupCode,
      name: r.clickupLabel || r.sourceLabel,
      itemSyncKey: itemSyncKey,
      dueDateRaw: dueValue,
      doneDateRaw: doneValue,
      note: noteValue,
      link: linkValue,
      description: description,
      dueCellA1Relative: toRelativeCellA1_(r.dueCol, r.dueRowOffset),
      doneCellA1Relative: toRelativeCellA1_(r.doneCol, r.doneRowOffset),
      noteCellA1Relative: toRelativeCellA1_(r.noteCol, r.noteRowOffset),
      linkCellA1Relative: toRelativeCellA1_(r.linkCol, r.linkRowOffset)
    });
  });

  return payloads;
}

function toRelativeCellA1_(col, rowOffset) {
  var c = stringValue_(col);
  if (!c || rowOffset === null || rowOffset === undefined || rowOffset === '') return '';
  return c.toUpperCase() + rowOffset + '(relative)';
}

function buildSubtaskDescription_(meta) {
  var lines = [];
  lines.push('# 업무 요약');
  lines.push('- 예정일: ' + formatDateYmd_(meta.dueValue, meta.tz));
  lines.push('- 완료일: ' + formatDateYmd_(meta.doneValue, meta.tz));
  lines.push('- 관련 메모: ' + (meta.noteValue || ''));
  lines.push('- 관련 링크: ' + (meta.linkValue || ''));
  lines.push('');
  lines.push('[SYNC_META]');
  lines.push('project_unique_id=' + meta.projectUniqueId);
  lines.push('item_sync_key=' + meta.itemSyncKey);
  lines.push('source_group=' + (meta.groupCode || ''));
  lines.push('source_item_name=' + (meta.sourceLabel || ''));
  lines.push('source_due_cell=' + (meta.dueCell || ''));
  lines.push('source_done_cell=' + (meta.doneCell || ''));
  lines.push('source_note_cell=' + (meta.noteCell || ''));
  lines.push('source_link_cell=' + (meta.linkCell || ''));
  lines.push('[/SYNC_META]');
  return lines.join('\n');
}
