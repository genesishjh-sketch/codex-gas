/** ClickUpConstants.gs
 * ClickUp 연동 기본 상수/기본 설정값.
 */

var CLICKUP_SETTINGS = {
  SHEET_NAME: 'clickup settings',
  LOG_SHEET_NAME: 'sync_logs',

  BASIC_CONFIG_DEFAULTS: [
    ['SOURCE_SHEET_NAME', '통합관리시트', '통합 원본 시트명'],
    ['SETTINGS_SHEET_NAME', 'clickup settings', '설정 시트명'],
    ['BLOCK_START_ROW', '4', '프로젝트 블록 시작행'],
    ['BLOCK_HEIGHT', '9', '프로젝트 블록 높이'],
    ['REPRESENTATIVE_ROW_OFFSET', '1', '대표행 오프셋(startRow 기준)'],
    ['CLICKUP_CREATE_STATUS_COLUMN', 'AC', 'clickup_create_status 컬럼'],
    ['PROJECT_UNIQUE_ID_COLUMN', 'X', 'project_unique_id 컬럼'],
    ['HOME_ITEM_KEY_COLUMN', 'Y', '홈스타일 item_sync_key 컬럼'],
    ['SUPPORT_ITEM_KEY_COLUMN', 'Z', '지원일정 item_sync_key 컬럼'],
    ['CLICKUP_TASK_ID_COLUMN', 'U', 'clickup_task_id 컬럼'],
    ['CLICKUP_TASK_URL_ROW_OFFSET', '1', '대표행 대비 task_url 행 오프셋'],
    ['SYNC_STATUS_ROW_OFFSET', '2', '대표행 대비 sync_status 행 오프셋'],
    ['LAST_SYNCED_AT_ROW_OFFSET', '3', '대표행 대비 last_synced_at 행 오프셋'],
    ['CREATE_CHECK_ROW_OFFSET', '4', '대표행 대비 생성체크 행 오프셋']
  ],

  CLICKUP_AUTH_DEFAULTS: [
    ['CLICKUP_API_TOKEN', '', 'ClickUp API 토큰 (필수)'],
    ['CLICKUP_EMAIL', '', 'ClickUp 계정 이메일 (선택/참고)']
  ],

  CLICKUP_TARGET_DEFAULTS: [
    ['CLICKUP_WORKSPACE_ID', '', 'ClickUp Workspace ID (선택)'],
    ['CLICKUP_SPACE_ID', '', 'ClickUp Space ID (선택)'],
    ['CLICKUP_FOLDER_ID', '', 'ClickUp Folder ID (선택)'],
    ['CLICKUP_LIST_ID', '', 'ClickUp List ID (필수)'],
    ['CLICKUP_DEFAULT_PARENT_STATUS', '', '부모 Task 기본 상태명 (선택)'],
    ['CLICKUP_DEFAULT_SUBTASK_STATUS', '', 'Subtask 기본 상태명 (선택)'],
    ['CLICKUP_COMPLETED_STATUS', '', '완료 처리 상태명 (선택)'],
    ['CLICKUP_ENABLE_DUE_DATE', 'TRUE', 'due date 활성화 TRUE/FALSE'],
    ['CLICKUP_PARENT_TITLE_INCLUDE_PROJECT_ID', 'FALSE', '부모 Task 제목에 project_unique_id 포함 여부(TRUE/FALSE)']
  ],

  ITEM_MAP_HEADERS: [
    'use_yn', 'item_code', 'group_code', 'source_label', 'clickup_label',
    'due_col', 'due_row_offset', 'done_col', 'done_row_offset',
    'note_col', 'note_row_offset', 'link_col', 'link_row_offset',
    'sort_order', 'description'
  ],

  ITEM_MAP_DEFAULT_ROWS: [
    ['Y', 'HOME_MEASURE', 'HOME', '실측', '실측', 'H', '1', 'I', '1', '', '', 'K', '0', '1', '홈스타일링 실측'],
    ['Y', 'HOME_CONSULT', 'HOME', '상담', '상담', 'H', '2', 'I', '2', 'G', '7', '', '', '2', '홈스타일링 상담'],
    ['Y', 'HOME_DESIGN', 'HOME', '디자인', '디자인', 'H', '3', 'I', '3', '', '', 'K', '3', '3', '홈스타일링 디자인'],
    ['Y', 'HOME_SHOPPING_LIST', 'HOME', '구매리스트', '구매리스트', 'H', '4', 'I', '4', '', '', 'K', '4', '4', '홈스타일링 구매리스트'],
    ['Y', 'HOME_SETTING', 'HOME', '세팅', '세팅', 'H', '5', 'I', '5', 'F', '7', '', '', '5', '홈스타일링 세팅'],
    ['Y', 'SUPPORT_CONSTRUCTION_DONE', 'SUPPORT', '시공종료', '시공종료', 'N', '1', 'O', '1', 'M', '7', '', '', '6', '지원 일정 시공종료'],
    ['Y', 'SUPPORT_FLORAL_DRAFT', 'SUPPORT', '조화시안', '조화시안', 'N', '2', 'O', '2', 'M', '7', '', '', '7', '지원 일정 조화시안'],
    ['Y', 'SUPPORT_FIRST_SETTING', 'SUPPORT', '1차 세팅일', '1차 세팅일', 'N', '3', 'O', '3', '', '', '', '', '8', '지원 일정 1차 세팅일'],
    ['Y', 'SUPPORT_FINAL_SETTING', 'SUPPORT', '최종세팅일', '최종세팅일', 'N', '4', 'O', '4', '', '', '', '', '9', '지원 일정 최종세팅일'],
    ['Y', 'SUPPORT_FLORAL_SCHEDULE', 'SUPPORT', '조화일정', '조화일정', 'N', '5', 'O', '5', '', '', '', '', '10', '지원 일정 조화일정']
  ]
};
