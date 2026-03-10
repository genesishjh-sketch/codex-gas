/** Todoist 연동 상수/공통 정의 */
var TODOIST_SYNC = {
  SETTINGS_SHEET_NAME: 'settings',
  DEFAULT_TARGET_SHEET: 'milestones',
  PROPERTY_API_TOKEN: 'TODOIST_API_TOKEN',
  // 2025-10 이후 구형 v8 엔드포인트가 완전 종료되어 /api/v1 기준으로 통일
  TODOIST_API_BASE_URL: 'https://api.todoist.com/api/v1',
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
  FULL_SYNC: {
    // Apps Script 실행 시간 제한(약 6분) 전에 안전하게 종료하기 위한 예산
    MAX_RUNTIME_MS: 330000,
    CURSOR_PROPERTY_KEY: 'TODOIST_FULL_SYNC_LAST_ROW'
  },
  INSTALLABLE_EDIT_TRIGGER_HANDLER: 'onMilestonesEditInstallable',
  DAILY_TRIGGER_HANDLER: 'runTodoistMilestonesFullSyncByTrigger'
};

var TODOIST_SETTINGS_LAYOUT = {
  columns: ['항목코드', '값', '설명', '예시'],
  sections: [
    {
      id: 'basic',
      type: 'keyValue',
      title: '기본 설정',
      rows: [
        {
          key: 'todoist_api_token',
          defaultValue: '',
          description: 'Todoist API 토큰(권장: settings 시트에 입력)',
          example: 'dummytoken1234567890'
        },
        {
          key: 'todoist_project_id',
          defaultValue: '',
          description: 'Todoist에서 작업을 생성할 프로젝트 ID',
          example: '1234567890'
        },
        {
          key: 'sync_target_sheet',
          defaultValue: 'milestones',
          description: '원본 데이터가 있는 시트명',
          example: 'milestones'
        },
        {
          key: 'due_date_field',
          defaultValue: 'plan_date',
          description: 'Todoist due date로 사용할 필드',
          example: 'plan_date'
        },
        {
          key: 'task_title_template',
          defaultValue: 'project_name&" | "&step_name&" 예정"',
          description: 'Todoist 할 일 제목 템플릿',
          example: "'TEXT(plan_date,\"mm-dd\")&\" \"&project_name&\" \"&step_name"
        },
        {
          key: 'label_template',
          defaultValue: '',
          description: 'Todoist 라벨 템플릿. 비워두면 라벨 미사용',
          example: 'section&"_"&step_name'
        },
        {
          key: 'exclude_done',
          defaultValue: true,
          description: 'done_date가 있는 행은 Todoist로 보내지 않음',
          example: 'TRUE'
        },
        {
          key: 'realtime_sync',
          defaultValue: true,
          description: '셀 입력/수정 시 바로 Todoist로 전송',
          example: 'TRUE'
        },
        {
          key: 'use_assignee',
          defaultValue: true,
          description: 'manager 값 기준으로 Todoist 담당자 배정(우선 ID, 없으면 이메일로 협업자 조회)',
          example: 'TRUE'
        },
        {
          key: 'use_description',
          defaultValue: false,
          description: '설명(description) 사용 여부',
          example: 'FALSE'
        },
        {
          key: 'use_labels',
          defaultValue: false,
          description: '라벨 기능 사용 여부',
          example: 'FALSE'
        }
      ]
    },
    {
      id: 'sectionMapping',
      type: 'table',
      title: '섹션 매핑',
      header: ['section값', 'todoist_section_id', '설명', '예시'],
      rows: [
        ['', '', 'B열 section 값과 Todoist 섹션 ID를 매핑', '']
      ]
    },
    {
      id: 'managerMapping',
      type: 'table',
      title: '담당자 매핑',
      header: ['manager_name', 'todoist_user_email', 'todoist_user_id', 'active', '설명'],
      rows: [
        ['', '', '', 'TRUE', 'manager_name 일치. user_id가 비어 있으면 이메일로 협업자에서 자동 매핑']
      ]
    }
  ]
};
