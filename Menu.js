/** 사용자 메뉴 등록 전용 */
function addInteriorSyncMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('🛋️ 인테리어 관리')
    .addItem('DB 동기화 실행', 'runInteriorDbSync')
    .addSeparator()
    .addItem('열 때 자동 동기화 켜기', 'enableInteriorSyncOnOpen')
    .addItem('열 때 자동 동기화 끄기', 'disableInteriorSyncOnOpen')
    .addSeparator()
    .addItem('Todoist 동기화 설치', 'setupTodoistMilestonesSync')
    .addItem('Todoist 현재 행 동기화', 'syncSelectedMilestoneRowToTodoist')
    .addItem('Todoist 전체 재동기화', 'runTodoistMilestonesFullSync')
    .addItem('Todoist 완료 반영(1시간) 설치', 'installHourlyTodoistCompletionMirrorTrigger')
    .addItem('Todoist 완료 반영(1시간) 제거', 'removeHourlyTodoistCompletionMirrorTriggers')
    .addItem('Todoist 매일 오전 자동동기화 설치', 'installDailyTodoistSyncTrigger9am')
    .addItem('Todoist 매일 자동동기화 제거', 'removeDailyTodoistSyncTriggers')
    .addSeparator()
    .addItem('설정 기준 자동 동기화 설치', 'installDailyInteriorSyncTriggerBySettings')
    .addItem('매일 오전 6시 자동 동기화 설치', 'installDailyInteriorSyncTrigger6am')
    .addItem('실시간 동기화 설치(변경행만)', 'installRealtimeInteriorSyncTrigger')
    .addItem('실시간 동기화 제거', 'removeRealtimeInteriorSyncTriggers')
    .addItem('매일 자동 동기화 제거', 'removeDailyInteriorSyncTriggers')
    .addToUi();
}
