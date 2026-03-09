# Codex GAS 스크립트

## 운영 주의사항
- **ContactsApp는 종료(Deprecated)** 되었습니다. 해당 환경에서는 연락처 동기화/점검이 자동으로 스킵되며, People API로 이전이 필요합니다.
- **People API 비활성화/권한 부족 시** 연락처 동기화·점검이 중단될 수 있으니 고급 서비스 및 Google Cloud Console 설정을 확인하세요.
- **Drive 체크는 DriveApp 권한이 필요**합니다. 권한이 없으면 드라이브 체크가 즉시 중단됩니다.

## 머지 충돌 시 필수 확인
- `<<<<<<<`, `=======`, `>>>>>>>` 표시가 남아 있으면 **스크립트가 실행되지 않습니다.**
- 반드시 충돌 마커를 제거하고, `ContactManager.js`의 연락처 헬퍼/동기화/점검 함수가 **정상 형태로 유지**되는지 확인하세요.

## 인테리어 DB 연동 프로그램 사용 설명서

### 1) 모듈 구조(유지보수 기준)
- `Menu.js`: 사용자 메뉴 등록
- `SyncService.js`: DB 동기화 핵심 로직 (`runInteriorDbSync` 등)
- `InteriorSettingsManager.js`: Script Properties 기반 설정 읽기/정규화
- `TriggerService.js`: 자동 실행 트리거 설치/제거
- `DashboardService.js`: 90일 KPI 집계 갱신
- `ArchiveService.js`: 완료 후 경과 마일스톤 보관 이관
- `AlertService.js`: 텔레그램 브리핑/지연 알림 발송

### 2) 최초 설정 순서
1. 스프레드시트 열기 → 메뉴 `🛋️ 인테리어 관리` 진입
2. Apps Script → 프로젝트 설정 → Script properties에 아래 키를 등록
   - `DAILY_SYNC_TIME_KST`: 동기화 시간(예: `08:30`, 24시간 형식)
   - `SYNC_SCOPE_MODE`: `지연만 / 7일예정만 / 지연+7일예정 / 전체`
   - `ARCHIVE_AFTER_DAYS`: 보관 이관 기준 일수(예: `30`)
3. 미등록 시 기본값(08:30, 지연+7일예정, 30일)이 사용됩니다.

### Todoist 설정 탭 안내
- `settings` 시트의 `todoist_api_token` 값을 우선 사용합니다.
- `todoist_api_token`이 비어 있으면 Script Properties의 `TODOIST_API_TOKEN`을 사용합니다.

### 3) 자동 동기화 설치
- 설정값 기준 자동 실행: `설정 기준 자동 동기화 설치`
- 고정 오전 6시 자동 실행: `매일 오전 6시 자동 동기화 설치`
- 자동 실행 제거: `매일 자동 동기화 제거`

### 4) 수동 실행/운영 함수
- 수동 DB 동기화: `runInteriorDbSync`
- 90일 대시보드 갱신: `refreshInteriorDashboard90d`
- 완료건 아카이브 이관: `archiveCompletedMilestones`
- 텔레그램 일일 브리핑: `sendTelegramDailyBriefing`
- 텔레그램 지연 알림: `sendTelegramDelayAlerts`

### 5) 텔레그램 알림 설정
`AlertService.js`는 Script Properties에서 아래 키를 읽습니다.
- `INTERIOR_TELEGRAM_BOT_TOKEN`
- `INTERIOR_TELEGRAM_CHAT_ID`

### 6) 트러블슈팅
- 메뉴가 안 보이면: 스프레드시트 새로고침 후 다시 열기
- 동기화가 실행되지 않으면: 트리거 중복 여부 확인 후 `매일 자동 동기화 제거` → 재설치
- 설정값 오류 시: Script properties 값을 다시 확인하고 저장하세요.
