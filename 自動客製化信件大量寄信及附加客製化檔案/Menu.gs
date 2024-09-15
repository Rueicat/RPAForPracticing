function onOpen() {
  const menu = SpreadsheetApp.getUi().createMenu('職護專用')
     menu
       .addItem('取得資料夾資料', 'getdata')
       .addSeparator()
       .addItem('大量寄信', 'mergeEmail')
       .addToUi();
}
