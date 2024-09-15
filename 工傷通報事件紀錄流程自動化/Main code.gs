////Globle Variable
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();
var sheetlastrow = sheet.getLastRow();
var data = sheet.getRange(2, 1, sheetlastrow, 22).getValues();

var 通報日期;
var 發生日期;
var 員工編號;
var 員工姓名;
var 通報類型;                                    
var 店鋪代號;                                        /////  <span><?= 通報內容 ?></span>    ////嵌入魔法
var 單位名稱;
var 部課;
var 職別;
var 通報內容;

var 標題;
var 已寄信件;
var 寄信店鋪信箱;
/////////////////////////////////////////

function myFunction() {
const confirm = Browser.msgBox('確認寄通知信?', Browser.Buttons.YES_NO);
const sentEmailsInfo = [];
if (confirm == 'yes') {
  for(let i = 0; i < data.length; i++ ){
    if(data[i][4] !== '' && data[i][21] !== true){       ////員編欄位不為空白  且 checkbox 沒有勾選  
      var rowData = data[i];
        let v通報日期 = rowData[1];
        通報日期 = Utilities.formatDate(v通報日期, "GMT+08:00", "yyyy/MM/dd");
        let v發生日期 = rowData[2];
        發生日期 = Utilities.formatDate(v發生日期, "GMT+08:00", "yyyy/MM/dd");
        員工編號 = rowData[4];
        員工姓名 = rowData[5];
        通報類型 = rowData[7];
        店鋪代號 = rowData[8];
        
        寄信店鋪信箱 = 店鋪代號.replace("B", "編號") + "@XXX.com.tw";

        單位名稱 = rowData[9];
        部課 = rowData[10];
        職別 = rowData[11];
        通報內容 = rowData[15];
        標題 = 通報類型 + "-" + 員工編號;
        已寄信件 = 員工姓名 + "---" + 通報類型 + "_" + 店鋪代號 +  "_" + 單位名稱;
        double_vlookup();
         sendEmail();

//////////////////////////////////////3/18加的, 還沒測試
   SpreadsheetApp.getActiveSpreadsheet().toast("已發送: " + 已寄信件, "通知", 5);
////////////////////////////////////////////////

       sheet.getRange(i + 2, 22).setValue(true);
        
         
        sentEmailsInfo.push(已寄信件);
    }else{}
       
       }                             /////////////////for circle
     if (sentEmailsInfo.length > 0) {
      var htmlOutput = HtmlService
      .createHtmlOutput('已寄出信件:<br>' + sentEmailsInfo.join('<br>'))
      .setWidth(400) // set width
      .setHeight(300); // set highth
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '寄信名單');
      
    } else {
      
      ss.toast('沒有需要寄出的信件', '💡提醒小視窗', 5);
    }
     

    
  }else{}
}