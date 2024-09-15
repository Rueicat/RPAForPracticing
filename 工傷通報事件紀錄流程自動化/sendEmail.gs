/*
1.利用google服務的特性, 可以讀取使用者的信箱, switch 該使用者的簽名檔
*/

function sendEmail(){

////定簽名檔
var 銳銳signature = HtmlService.createTemplateFromFile('銳銳signature').evaluate().getContent();
var osh_signature = HtmlService.createTemplateFromFile('osh_signature').evaluate().getContent();

///陣列儲存
var usersignature = {
"user@XXX.com.tw": 銳銳signature,
"osh@XXX.com.tw": osh_signature
}; 

///抓取登入者身分, 擷取信箱
var currentuser = Session.getActiveUser().getEmail();  
//利用他的信箱判斷用哪個客製化的簽名檔
var userProfile = usersignature[currentuser];   


var emailBody = HtmlService.createTemplateFromFile('template').evaluate().getContent();

emailBody += userProfile;

MailApp.sendEmail({

to: 部信箱 + "," + 區信箱 + "," + 寄信店鋪信箱 + ",osh@XXX.com.tw",
subject: 標題,
htmlBody: emailBody,
})

}