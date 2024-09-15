/*
1.創建客製化選項(在spreadsheet最上方那一排選項)
*/


function onOpen() {
  var menu = SpreadsheetApp.getUi().createMenu('職護的秘密');
      menu
      .addItem('寄出通知信', 'myFunction')
      .addToUi(); 
}
