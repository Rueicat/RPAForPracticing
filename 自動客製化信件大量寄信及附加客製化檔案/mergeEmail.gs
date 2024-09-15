function mergeEmail() {
  const confirm = Browser.msgBox('確認要開始-大量-寄信了?', Browser.Buttons.YES_NO);
  if (confirm == 'yes') {

  let emailbody = HtmlService.createTemplateFromFile("EmailBody").evaluate().getContent();
  let ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('信件區');
  let lastrow = ss.getLastRow();
  let datarange = ss.getRange(2, 1, lastrow -1, 6);
  let data = datarange.getValues();
  let completedData = []; // 設一個陣列容器


   ///神人用法, 可能看不懂
  for (let i in data){
    let row = data[i];
    let 店鋪信箱 = row[0];
    let 附件ID = row[4];
    let 標題店鋪檔名 = row[5];
    let 附件 = DriveApp.getFileById(附件ID);

    MailApp.sendEmail({
      to: 店鋪信箱 + ",osh@mos.com.tw",  
      subject: "XXXX年度在職夥伴體檢報告--" + 標題店鋪檔名,
      htmlBody: emailbody,
      attachments: [附件.getAs(MimeType.PDF)]
    });

    // 添加成功發送的訊息到陣列中
    completedData.push(標題店鋪檔名);
    
    
    SpreadsheetApp.getActiveSpreadsheet().toast("已發送: " + 標題店鋪檔名, "發送成功", 5);
  }



  // 創建並顯示完成資料的彈出窗口, this is function inner function mergeEmail()
  showCompletedDataPopup(completedData);
}else{}

}



function showCompletedDataPopup(completedData) {
  let htmlContent = "<html><head><title>發送完成</title></head><body><h2>以下是已完成發送的資料：</h2><ul>";
  completedData.forEach(function(dataItem) {
    htmlContent += "<li>" + dataItem + "</li>";
  });
  htmlContent += "</ul></body></html>";

  let html = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, "發送報告");
}
