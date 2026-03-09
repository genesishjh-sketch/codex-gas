/** Todoist 연동 상수/공통 정의 */
var TODOIST_SYNC = {
  SETTINGS_SHEET_NAME: 'settings',
  DEFAULT_TARGET_SHEET: 'milestones',
  PROPERTY_API_TOKEN: 'TODOIST_API_TOKEN',
  TODOIST_API_BASE_URL: 'https://api.todoist.com/rest/v2',
  MILESTONE_HEADERS: ['project_code', 'section', 'step_name', 'plan_date', 'done_date', 'manager'],
  SYNC_HEADERS: ['todoist_task_id', 'sync_status', 'last_synced_at', 'last_error'],
  STATUS: {
    CREATED: '전송완료',
    UPDATED: '업데이트완료',
    SKIPPED: '스킵',
    ERROR: '오류'
  },
  ASSIGNEE_POLICY: {
    // manager 매핑이 없을 때 동작: false면 assignee 없이 생성, true면 오류 처리
    ERROR_IF_NOT_FOUND: false
  },

  PROCESS_MARK: {
    COLUMN_INDEX: 12, // L열
    // 아래 텍스트가 포함되면 전체 재동기화에서 이미 처리된 행으로 간주하여 스킵
    SKIP_KEYWORDS: ['체크', '동기화완료', '투두이시트완료', 'todoist완료']
  },
  INSTALLABLE_EDIT_TRIGGER_HANDLER: 'onMilestonesEditInstallable',
  DAILY_TRIGGER_HANDLER: 'runTodoistMilestonesFullSyncByTrigger'
};
