///Globle variety
////<span><?= 安插變數 ?></span>
////2024.7.31 修復表格問題
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sss = ss.getSheetByName('直式通知單')


var 店鋪標題 = sss.getRange('B1').getValues();
var 月 = sss.getRange('I4').getValue();
var 日 = sss.getRange('J4').getValue();
var 時 = sss.getRange('K4').getValue();
var 分 = sss.getRange('L4').getValue();

var 體檢一 = sss.getRange('C15').getValue();
var 體檢二 = sss.getRange('C16').getValue();
var 體檢三 = sss.getRange('C17').getValue();
var 體檢四 = sss.getRange('C18').getValue();
var 體檢五 = sss.getRange('C19').getValue();
var 體檢六 = sss.getRange('C20').getValue();
var 體檢七 = sss.getRange('C21').getValue();
var 體檢八 = sss.getRange('C22').getValue();
var 體檢九 = sss.getRange('C23').getValue();
var 體檢十 = sss.getRange('C24').getValue();
var 體檢十一 = sss.getRange('C25').getValue();
var 體檢十二 = sss.getRange('C26').getValue();
var 體檢十三 = sss.getRange('C27').getValue();
var 體檢十四 = sss.getRange('C28').getValue();
//////
var 負荷一 = sss.getRange('F15').getValue();
var 負荷二 = sss.getRange('F16').getValue();
var 負荷三 = sss.getRange('F17').getValue();
var 負荷四 = sss.getRange('F18').getValue();
var 負荷五 = sss.getRange('F19').getValue();
var 負荷六 = sss.getRange('F20').getValue();
var 負荷七 = sss.getRange('F21').getValue();
var 負荷八 = sss.getRange('F22').getValue();
var 負荷九 = sss.getRange('F23').getValue();
var 負荷十 = sss.getRange('F24').getValue();
var 負荷十一 = sss.getRange('F25').getValue();
var 負荷十二 = sss.getRange('F26').getValue();
var 負荷十三 = sss.getRange('F27').getValue();
var 負荷十四 = sss.getRange('F28').getValue();
//////
var 肌肉一 = sss.getRange('I15').getValue();
var 肌肉二 = sss.getRange('I16').getValue();
var 肌肉三 = sss.getRange('I17').getValue();
var 肌肉四 = sss.getRange('I18').getValue();
var 肌肉五 = sss.getRange('I19').getValue();
var 肌肉六 = sss.getRange('I20').getValue();
var 肌肉七 = sss.getRange('I21').getValue();
var 肌肉八 = sss.getRange('I22').getValue();
var 肌肉九 = sss.getRange('I23').getValue();
var 肌肉十 = sss.getRange('I24').getValue();
var 肌肉十一 = sss.getRange('I25').getValue();
var 肌肉十二 = sss.getRange('I26').getValue();
var 肌肉十三 = sss.getRange('I27').getValue();
var 肌肉十四 = sss.getRange('I28').getValue();

var 高齡一 = sss.getRange('L15').getValue();
var 高齡二 = sss.getRange('L16').getValue();
var 高齡三 = sss.getRange('L17').getValue();
var 高齡四 = sss.getRange('L18').getValue();
var 高齡五 = sss.getRange('L19').getValue();
var 高齡六 = sss.getRange('L20').getValue();
var 高齡七 = sss.getRange('L21').getValue();
var 高齡八 = sss.getRange('L22').getValue();
var 高齡九 = sss.getRange('L23').getValue();
var 高齡十 = sss.getRange('L24').getValue();
var 高齡十一 = sss.getRange('L25').getValue();
var 高齡十二 = sss.getRange('L26').getValue();
var 高齡十三 = sss.getRange('L27').getValue();
var 高齡十四 = sss.getRange('L28').getValue();

/////////////////////



function send_mail() {
const confirm = Browser.msgBox('確認寄出信件?', Browser.Buttons.YES_NO);
if(confirm == 'yes') {
  



double_vlookup();

var ssheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('臨場服務明細表');
  var Bshop = ssheet.getRange('Q2').getValue();
var Bname = ssheet.getRange('S2').getValue();
var PDFName = Bshop + Bname + "- 臨場服務通知單.pdf";


var 寄信店鋪信箱 = Bshop.replace("B", "number") + "@XXX.com.tw";


  var emailBody = HtmlService.createTemplateFromFile('template').evaluate().getContent();

  var blob = Utilities.newBlob(emailBody, 'text/html').getAs('application/pdf').setName(PDFName);

  
MailApp.sendEmail({
to: 部信箱 + "," + 區信箱 + "," + 寄信店鋪信箱 + ",XXXX@XXX.com.tw",
subject: "XXXX年" + 月 + 日 + "-" + Bshop + Bname + "-臨場服務通知",
htmlBody: emailBody,
attachments: [blob],
})



ss.toast('信件已經寄出', '💡提醒小視窗', 5);
}
}
