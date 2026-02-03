/** Config.gs */

/** [설정] 카카오 REST API 키 */
const KAKAO_API_KEY = 'a5949a3368f9a0d1de67758e41bdd158';

/**
 * [설정] 시트/블록 구조
 * - 업로드된 엑셀 기준: 9행 1블록 (4,13,22,...)
 */
const CONFIG = {
  SHEET_NAME: "통합관리시트",
  REPORT_SHEET_NAME: "업무보고",
  CALENDAR_SHEET_NAME: "홈스타일_주간일정",

  /** ✅ DB 시트명: 비우면 자동 생성(권장) */
  DB_SHEET_NAME: "",

  START_ROW: 4,
  BLOCK_HEIGHT: 9,

  /** ✅ 종료 조건: (B,C) 모두 빈 블록이 연속 N개면 중단 */
  STOP_AFTER_EMPTY_BLOCKS: 3,

  /** 데이터 위치 (1-based) */
  POS_NO:     { row: 0, col: 2 },   // B: No
  POS_NAME:   { row: 0, col: 3 },   // C: 프로젝트명(수식)

  POS_ADDR:   { row: 0, col: 6 },   // F4: 주소(변환 결과 저장)
  POS_ADDR_EXTRA: { row: 2, col: 6 }, // F6: 지번 뒤 추가정보(호/층 등)
  POS_MAP:    { row: 4, col: 6 },   // F8: 지도 URL

  /**
   * ✅ 물품리스트 링크 위치
   * - H 오른쪽에 2열 추가한 구조라면 I→K로 밀렸을 가능성이 큼
   * - 네 시트에서 “물품리스트 링크 셀”이 실제로 K8이면 col:11 이 맞음
   */
  POS_FILE:   { row: 4, col: 11 },  // K8: 물품리스트 파일 링크(권장)

  /** 템플릿(물품리스트) 위치 */
  CELL_TEMPLATE_ORIGIN: "G1",

  /** 삭제예정 플래그(메모행(START+7) G열에 1이면 삭제예정) */
  DELETE_FLAG: { row: 7, col: 7 },

  /** 연락처 생성 옵션 */
  CONTACT_SKIP_IF_NO_PHONE: true,
  /** 연락처_log에 존재하면 스킵 (false면 재동기화 시도) */
  CONTACT_SKIP_IF_LOGGED: true,
  /** 연락처_log에 있어도 실제 연락처 존재 여부 검증 */
  CONTACT_VERIFY_LOGGED: true,
  /** 스킵 사유를 연락처_log에 남길지 여부 */
  CONTACT_LOG_SKIP_REASONS: true,
  /** 프로젝트명 유효성 검사 무시하고 연락처 동기화 */
  CONTACT_IGNORE_NAME_VALIDATION: false,

  /** ✅ 프로젝트명 판정 옵션 */
  NAME_VALIDATION: {
    /** 접두어 목록(비어 있으면 접두어 강제하지 않음) */
    prefixes: [],
    /** "님"과 같은 접미어 강제 여부 */
    requireSuffix: false,
    /** 접미어 기본값 */
    suffix: "님",
    /** 접두어/접미어 무시하고 항상 유효로 처리 */
    allowAny: false
  },

  /** ✅ (카톡 보고서/DB) 홈스타일링 전용 컬럼: 라벨 G / 예정 H / 완료 I */
  HOME_TASK: {
    labelCol: 7, // G
    planCol: 8,  // H (예정일)
    doneCol: 9   // I (완료일)
  },

  /** (캘린더 등) 기존 호환용 - 필요시만 사용 */
  TASK_COLS: [
    { labelCol: 7,  dateCol: 8,  prefix: "" },        // 홈스타일링 예정일(H)
    { labelCol: 13, dateCol: 14, prefix: "[시공] " }  // (H 오른쪽 2열 추가 후) 시공 M/N 가정
  ],

  /** ✅ 드라이브 체크(캐시/표시) */
  DRIVE_LOG_SHEET: "드라이브_check_log",
  DRIVE_CACHE_HOURS: 6,
  DRIVE_MARK_COL: 19 // S열(표시용 색칠)
};
