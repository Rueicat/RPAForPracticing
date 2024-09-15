function clearWords() {
  
  const confirm = Browser.msgBox('確定清空體檢安排格?', Browser.Buttons.YES_NO);
if(confirm == 'yes') {
  
  
  // 獲取當前活動的試算表
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // 設定起始儲存格的行號
  const startRow = 6;
  
  // 清除指定儲存格及其後每 4 個儲存格的內容
  for (let i = startRow; i <= 186; i += 4) {
    const range = sheet.getRange("Q" + i);
    range.clearContent();
  }




ss.toast('清理完畢~', '💡提醒小視窗', 5);
  }
}
