/**
 * 當試算表開啟時，自動建立自訂選單
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // 建立一個選單
  ui.createMenu(' 新增&同步')
    .addItem('1. 建立下週會議表 (含目錄同步)', 'createNextMeetingSheet')
    .addItem('2. 同步 Trello 請假人員', 'syncLeaveStatus')
    .addSeparator() // 加入分隔線
    .addItem('說明：第一次執行需授權', 'showInstruction') 
    .addToUi();
}

/**
 * 點選說明的簡單提示框
 */
function showInstruction() {
  const ui = SpreadsheetApp.getUi();
  ui.alert("使用說明", "若點選功能沒反應，請先確認是否已完成 Google 權限授權。\n如有報錯，請聯繫腳本管理員。", ui.ButtonSet.OK);
}
