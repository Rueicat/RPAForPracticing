///update on 2024.8/14
////2024.5/17, Script by Ruei Kato
function doPost(e) {
  try{
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clients');
    var data = JSON.parse(e.postData.contents);
    var rows = sheet.getDataRange().getValues();
    switch (data.type){
      case 'register':
        for (var i = 0; i < rows.length; i ++){
          if (rows[i][0] === data.username){
              return ContentService.createTextOutput('已經註冊過了, 請選擇其他用戶或選擇其他功能');
          } else{
              sheet.appendRow([data.username, data.password, data.salt, data.useremail]);
              return ContentService.createTextOutput("註冊成功");
          }
        }                    

      case 'change':
      case 'validate':
        var userRow = rows.find(row => row[0] === data.username);
          if (userRow){
            if (userRow[1] === data.password){
              return ContentService.createTextOutput('密碼正確, 驗證成功');
            } else{
              return ContentService.createTextOutput('密碼錯誤, 驗證失敗');
            }
          } else {
            ///找鹽值時已經有確認是否有這帳號, 這邊不用判定
          }
      case 'getSalt':
        var userRow = rows.find(row => row[0] === data.username);
        if (userRow) {
          return ContentService.createTextOutput(userRow[2]);
        } else {
          return ContentService.createTextOutput('未找到用戶');
        }
    }
 } catch (error){
    return ContentService.createTextOutput("處理過程中發生錯誤: " + error.message);
 }
}




