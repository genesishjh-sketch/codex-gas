/** 사용자 메뉴 등록 전용 */
function addInteriorSyncMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('🛋️ 인테리어 관리')
    .addItem('DB 동기화 실행', 'runInteriorDbSync')
    .addSeparator()
    .addItem('열 때 자동 동기화 켜기', 'enableInteriorSyncOnOpen')
    .addItem('열 때 자동 동기화 끄기', 'disableInteriorSyncOnOpen')
    .addSeparator()
    .addItem('[설치] 설정 기준 자동 동기화', 'installDailyInteriorSyncTriggerBySettings')
    .addItem('[설치] 매일 오전 6시 자동 동기화', 'installDailyInteriorSyncTrigger6am')
    .addItem('[설치] 실시간 동기화(변경행만)', 'installRealtimeInteriorSyncTrigger')
    .addItem('[제거] 실시간 동기화', 'removeRealtimeInteriorSyncTriggers')
    .addItem('[제거] 매일 자동 동기화', 'removeDailyInteriorSyncTriggers')
    .addToUi();
}
