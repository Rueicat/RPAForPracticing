function uploadSignature(dataURL) {
  if (dataURL) {
    try {
      var base64String = dataURL.split(',')[1];
      var byteString = Utilities.base64Decode(base64String);
      var blob = Utilities.newBlob(byteString, 'image/png', 'signature.png');

      // 選擇Google Drive中的資料夾來儲存簽名圖像
      DriveApp.getFolderById('urspreadsheetID').createFile(blob);
    } catch (e) {
      Logger.log("Error in uploadSignature: " + e.toString());
      return "Error: " + e.toString();
    }
  } else {
    Logger.log("Data URL is undefined or not provided.");
    return "Error: Data URL is undefined or not provided.";
  }
}

// 部署為Web應用程序
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Page').setTitle('手機簽名板');
}
