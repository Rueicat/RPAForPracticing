//不知道為什麼, 轉換取值時(excel to spreadsheet), cell 後面有時候會多很多白, 這個要修正去掉(用trim)

function removeTrailingSpaces() {
  // 試算表ID
  var spreadsheetId = '欲處理的試算表ID位置';
  
  // 獲取試算表
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheets()[0];
  
  // 要處理的欄位
  var columns = ['D', 'E', 'F'];
  
  columns.forEach(function(column) {
    // 獲取該欄的所有數據
    var data = sheet.getRange(column + '1:' + column).getValues();
    
    // 去除每個儲存格後面的空白
    for (var i = 0; i < data.length; i++) {
      if (data[i][0]) {
        data[i][0] = data[i][0].toString().replace(/\s+$/, '');
      }
    }
    
    // 更新試算表中的數據
    sheet.getRange(column + '1:' + column).setValues(data);
  });
}

