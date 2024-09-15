/*
1.spreadsheet前端表格有設計排班, 但因為很多, 比對有無重複安排很花時間, 所以寫程式自動判定
*/

function check() {
  const confirm = Browser.msgBox('確認訊息, 是抓取V欄位判定, 確認跑程式?', Browser.Buttons.OK_CANCEL);

  if (confirm == 'ok'){
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('XXXX年總表');
      
      

      var column = sheet.getRange('V:V').getValues();
      var columnData = column.map(function(row) { return row[0];}).filter(value => value);    /// 變1' array     
      
      var seen = new Map();
      var duplicates = new Map();

      for (var i = 0; i < columnData.length; i++){
          var value = columnData[i];
          if (seen.has(value)){
            duplicates.set(value, (duplicates.get(value) || 0) + 1);
          } else {
            seen.set(value, 1);            
          }
      }
         ////output
         var output = [];
         duplicates.forEach(function(count, value){
           output.push(value + '重複次數：' + (count + 1));
         });

         if (output.length > 0) {
           Browser.msgBox(output.join('\n'));
         } else {
            Browser.msgBox('沒有重複店舖安排')
         }
        
  } else {
    /////如果按取消, 目前沒有要幹嘛, 就取消

   }
}
