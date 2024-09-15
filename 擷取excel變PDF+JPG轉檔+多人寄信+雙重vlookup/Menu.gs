function onOpen() {
  const menu = SpreadsheetApp.getUi().createMenu('職護專用')
    menu 
      .addItem('自動寄通知信', 'sendinform')
      .addToUi();
}

