/** 사용자 메뉴 등록 전용 */
function addInteriorSyncMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('🛋️ 인테리어 관리')
    .addItem('DB 동기화 실행', 'runInteriorDbSync')
    .addSeparator()
    .addItem('열 때 자동 동기화 켜기', 'enableInteriorSyncOnOpen')
    .addItem('열 때 자동 동기화 끄기', 'disableInteriorSyncOnOpen')
    .addSeparator()
    .addItem('설정 기준 자동 동기화 설치', 'installDailyInteriorSyncTriggerBySettings')
    .addItem('매일 오전 6시 자동 동기화 설치', 'installDailyInteriorSyncTrigger6am')
    .addItem('매일 자동 동기화 제거', 'removeDailyInteriorSyncTriggers')
    .addToUi();
}
