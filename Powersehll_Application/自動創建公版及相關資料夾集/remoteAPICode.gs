

  function doPost(e) {
    try {
  var requestBody = JSON.parse(e.postData.contents);
  var storeCode1 = requestBody.storeCode1;
  var storeCode2 = requestBody.storeCode2;

  var sheet = SpreadsheetApp.openById('1D9hlryVrvd1KUhHxU27shCnKOHUqTW7wE2bkIF8Vm4E').getSheets()[0];
  var data = sheet.getRange(1, 1,sheet.getLastRow(), sheet.getLastColumn()).getValues();

  var storeName1 = null;
  var storeName2 = null;


 ////// 神之方法, 
 for (var i = 1; i < data.length; i++) {
    if (data[i][0] === storeCode1) {
      storeName1 = data[i][1];
      if (storeName2) break;
    }
    if (data[i][0] === storeCode2) {
      storeName2 = data[i][1];
      if (storeName1) break;
    }
  }
 //////////////////////文件處理部分
     var templateId = '公版ID';
  var tempfolder = DriveApp.getFolderById('所在資料夾ID'); 
  /* Copy doc temp folder, 請一定要記得換成google文件, 
  不能是.docx, 我沒發現這個錯誤卡了一個星期都在做白工...一直執行錯誤....
  google文件編輯好了以後, 要取得文件的blob, 轉成docx 生成的下載連結才可以真的在其他平台使用
  */
  // 先刪除 temp 之前的舊檔案
  var oldfiles = tempfolder.getFiles();
  while (oldfiles.hasNext()) {
    oldfiles.next().setTrashed(true);
  }

  // 母檔
  var originaldoc = DriveApp.getFileById(templateId);
  var processdoc1ID = originaldoc.makeCopy(storeName1 + "-臨場服務改善提醒單", tempfolder).getId();
  var processdoc2ID = originaldoc.makeCopy(storeName2 + "-臨場服務改善提醒單", tempfolder).getId();

 

  var doc1 = DocumentApp.openById(processdoc1ID);
  var doc2 = DocumentApp.openById(processdoc2ID);

  var body1 = doc1.getBody();
  var body2 = doc2.getBody();

  body1.replaceText('<ShopCode>', storeName1);
  body2.replaceText('<ShopCode>', storeName2);

  doc1.saveAndClose();
  doc2.saveAndClose();

   // Convert to DOCX
  var token = ScriptApp.getOAuthToken();
  var blb1 = convertToDocx(processdoc1ID, token);
  var blb2 = convertToDocx(processdoc2ID, token);

  var file1 = DriveApp.createFile(blb1.setName(storeName1 + '.docx'));
  var file2 = DriveApp.createFile(blb2.setName(storeName2 + '.docx'));

  file1.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);
  file2.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);

  var downloadurl1 = file1.getDownloadUrl();
  var downloadurl2 = file2.getDownloadUrl();

  var result = {
    storeName1: storeName1,
    storeName2: storeName2,
    downloadurl1: downloadurl1,
    downloadurl2: downloadurl2
  };


 
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }catch (error){
       return ContentService.createTextOutput(JSON.stringify({error: error.toString()})).setMimeType(ContentService.MimeType.JSON);

    }
}

function convertToDocx(docId, token) {
  return UrlFetchApp.fetch('https://docs.google.com/feeds/download/documents/export/Export?id=' + docId + '&exportFormat=docx', {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  }).getBlob();
}
