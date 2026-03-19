/** 사용자 메뉴 등록 전용 */
function addInteriorSyncMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('🛋️ 인테리어 관리')
    .addItem('DB 동기화 실행', 'runInteriorDbSync')
    .addItem('Todoist 현재 행 동기화', 'syncSelectedMilestoneRowToTodoist')
    .addItem('Todoist 전체 재동기화', 'runTodoistMilestonesFullSync')
    .addSeparator()
    .addItem('열 때 자동 동기화 켜기', 'enableInteriorSyncOnOpen')
    .addItem('열 때 자동 동기화 끄기', 'disableInteriorSyncOnOpen')
    .addSeparator()
    .addItem('[설치] Todoist 동기화 기본세팅', 'setupTodoistMilestonesSync')
    .addItem('[설치] Todoist 완료 반영(1시간)', 'installHourlyTodoistCompletionMirrorTrigger')
    .addItem('[제거] Todoist 완료 반영(1시간)', 'removeHourlyTodoistCompletionMirrorTriggers')
    .addItem('[설치] Todoist 매일 자동동기화(09시)', 'installDailyTodoistSyncTrigger9am')
    .addItem('[제거] Todoist 매일 자동동기화', 'removeDailyTodoistSyncTriggers')
    .addSeparator()
    .addItem('[설치] 설정 기준 자동 동기화', 'installDailyInteriorSyncTriggerBySettings')
    .addItem('[설치] 매일 오전 6시 자동 동기화', 'installDailyInteriorSyncTrigger6am')
    .addItem('[설치] 실시간 동기화(변경행만)', 'installRealtimeInteriorSyncTrigger')
    .addItem('[제거] 실시간 동기화', 'removeRealtimeInteriorSyncTriggers')
    .addItem('[제거] 매일 자동 동기화', 'removeDailyInteriorSyncTriggers')
    .addToUi();
}
