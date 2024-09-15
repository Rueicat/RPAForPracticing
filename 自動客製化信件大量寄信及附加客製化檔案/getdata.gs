function getdata() {
   var html = HtmlService.createHtmlOutputFromFile('Picker.html')
      .setWidth(600)
      .setHeight(425)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select Folder');
}

function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}


/*
Function that takes item Id from Picker.html once user has made selection.
Creates clickable Url in spreadsheet.
Pastes in item Id to spreadsheet.
*/

function insertFolderURL(id){

var folder = DriveApp.getFolderById(id);
var files = folder.getFiles();    ////All the files in this folder


///list 是java script的用法, list.push[1, 2, 3], list.push[4], 是一直添加到list變[1, 2, 3, 4], 但都是虛擬資料, 
///需要"輸出",  也可以用appendrow, 用法, 這個比較是專門給sheet用的, 添加到最後一行空白開始. 
var list = [];
list.push(['收件者(自動)', '附件檔名(自動)', '附件ID(自動)', '檔名(自動)']);

//// use regular expression


   while (files.hasNext()){
     var file = files.next();
     var filenoextension = file.getName().replace(/(.+)\.[^.]+$/, "$1");

     var matchBB = file.getName().match(/B(\d{3})/);
     var matchBBB = matchBB ? matchBB[1] : "000";

     var row = [];
     row.push("preupper" + matchBBB + "@XXX.com.tw", file.getName(), file.getId(), filenoextension);  ////因為id還沒抓到資料, 所以沒有東西, 暫時顯示錯誤
     list.push(row);
   }

var listdata = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('(勿動)阿波羅工作區');
var numrows = list.length;
var numcols = list[0].length;

 listdata.clearContents();

 listdata.getRange(1, 1,numrows, numcols).setValues(list);
  
}