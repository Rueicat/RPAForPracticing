//在試算表上方創建客製化選單執行程式

function onOpen() {
    var menu = SpreadsheetApp.getUi().createMenu('檢查重複');
       menu
       .addItem('檢查是否重複安排店鋪', 'check')
       .addToUi();
}
