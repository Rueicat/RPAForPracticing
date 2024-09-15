////Ver.5 是一般完整功能的版本
////Ver.7 是增加IP擷取功能
////Ver.8 是增加查詢IP位置地點的功能   2024.6.2
////Movement by Ruei
////記得要部屬

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index');
}

function searchId(idNumber) {
  var sheetId = '欲查詢表單的ID位置on google drive';
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('(原檔)XXXX年度體檢表單'); // Adjust the sheet name if necessary
  var data = sheet.getDataRange().getValues();

  var result = null;
  for (var i = 1; i < data.length; i++) {
    if (data[i][3] == idNumber) { // Assuming 身分證號 is in column D (index 3)
      result = {
        name: data[i][2], // Assuming 姓名 is in column C (index 2)
        clinic: data[i][0], // Assuming 體檢單位 is in column A (index 0)
        date: data[i][8] // Assuming 體檢日期 is in column I (index 8)
      };
      break;
    }
  }

  var ipAddress = getClientIp();
  var location = getIpLocation(ipAddress);

  logSearch(idNumber, result, ipAddress, location);

  return result;
}

function getClientIp() {
  var response = UrlFetchApp.fetch('https://api.ipify.org?format=json');
  var ip = JSON.parse(response.getContentText()).ip;
  return ip;
}

function getIpLocation(ipAddress) {
  var response = UrlFetchApp.fetch('https://ipinfo.io/' + ipAddress + '/json');
  var locationData = JSON.parse(response.getContentText());
  var location = locationData.city + ', ' + locationData.region + ', ' + locationData.country;
  return location;
}

function logSearch(idNumber, result, ipAddress, location) {
  var logSheetId = '1ynwQ1PnTVa-JwHyXCqWZVJG-EwMxzNX5MlORV8uabNQ';
  var logSheet = SpreadsheetApp.openById(logSheetId).getSheetByName('搜尋紀錄'); // Adjust the sheet name if necessary
  var timestamp = new Date();
  logSheet.appendRow([timestamp, idNumber, result ? result.name : '', result ? result.clinic : '', result ? result.date : '', ipAddress, location]);
}
