/**
 * 2024.8.13
 * 自動擷取LINE 通報群組PO的訊息,
 * 貼到自己指定的通報事件紀錄spreadsheet
 * 
 */


const LINE_Channel_Access_Token = 'urToken';

const ssID = 'WhereToPasteTheInformation_catchedFromLINEBOT';


///Webhook request from LINE
function doPost(e) {
  var json = JSON.parse(e.postData.contents);
  var events = json.events;

  events.forEach(function(event){
      if(event.type === 'message' && event.message.type === 'text'){
      const userMessage = event.message.text;
      extraAndSaveNumbers(userMessage);
      }
  });
}

function extraAndSaveNumbers(message){
    const regex = /\d{9}/g;     /////regular expression, extract ID numbers
    const matches = message.match(regex);
    
    if(matches){
    const sheet_Name = getCurrentMonthSheetName();
    const sheet = SpreadsheetApp.openById(ssID).getSheetByName(sheet_Name);
    const currentDate = new Date();
    const formatteDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
    const currentMonth = currentDate.getMonth() + 1;

        matches.forEach(function(number){
            /// 更新整個row sheet.appendRow([currentMonth, formatteDate, "", "", number, "", "", "", "", "", "","", "", "", "", message]);
          const lastRow = sheet.getRange('E:E').getValues().filter(String).length + 1;   ////filter 過濾空白
          sheet.getRange(lastRow, 1).setValue(currentMonth);
          sheet.getRange(lastRow, 3).setValue('請人工確認');
          sheet.getRange(lastRow, 2).setValue(formatteDate);
          sheet.getRange(lastRow, 5).setValue(number);
          sheet.getRange(lastRow, 16).setValue(message);
          sheet.getRange(lastRow, 19).setValue('未確認');
        })
    } else{
      const sheet_Name = getCurrentMonthSheetName();
      const sheet = SpreadsheetApp.openById(ssID).getSheetByName(sheet_Name);
       const lastRow = sheet.getRange('E:E').getValues().filter(String).length + 1;

       const currentDate = new Date();
       const formatteDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy/MM/dd");


       sheet.getRange(lastRow, 5).setValue('資料無法判定, 請人工確認:' + message);
       sheet.getRange(lastRow, 2).setValue(formatteDate);
    }
}


function getCurrentMonthSheetName(){
    /////以下array 是屬順序的, 第一個是0, 第二個是1...
    const monthNames = ["一月", "二月", "三月", "四月", "五月", "六月", "七月",
                   "八月", "九月", "十月", "十一月", "十二月"];    
      const currentdate = new Date();
      return monthNames[currentdate.getMonth()];

}

