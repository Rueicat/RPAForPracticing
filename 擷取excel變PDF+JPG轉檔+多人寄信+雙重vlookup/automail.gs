function autoemail() {

let image = DriveApp.getFileById(圖片的ID).getAs("image/png");
let mailImages = {"logo": image};
let emailBody = HtmlService.createTemplateFromFile('template').evaluate().getContent();

MailApp.sendEmail({
  to: 部信箱 + "," + 區信箱 + "," + 寄信店鋪信箱 + ",otheruser@XXX.com.tw",
  subject: 月 + 日 + "-" + 店鋪代號 + "-" + 店鋪名稱 + "臨場服務通知單",
  htmlBody: emailBody,
  inlineImages: mailImages,
  attachments: [要寄信的pdf]                ////////////記得加s
})
}
