function doPost(e) {
    var jsonString = Utilities.newBlob(Utilities.base64Decode(e.postData.contents)).getDataAsString();

    var PSdata = JSON.parse(jsonString);
    var data = PSdata.data;
    var fileName = PSdata.fileName;
    var leaveData = PSdata.leaveData;
    var Interval = PSdata.Interval;

    var folderID = '1o7rLbGjN8itvGnvFBJgmieSXyTTg_2cd';
    var folder = DriveApp.getFolderById(folderID);
    
    var ss = SpreadsheetApp.create(fileName);
    var ssID = ss.getId();
    var sheet = ss.getSheets()[0];
    var last4Char = fileName.slice(-4);
    sheet.setName(last4Char);

    //////處理第二個sheet 請假單的部分
      var sheet2 = ss.insertSheet('請假單');
          sheet2.getRange(1, 1, 1, 1).setValue('看原始資料寫什麼, 複寫過去');
          sheet2.getRange(2, 1, 1, 1).setValue('請假單明細表');
          sheet2.getRange(3,1, 1, 1).setValue(Interval);

          ////模仿人事資料的map方法倒入資料
          var LLheaders = Object.keys(leaveData[0]);
          sheet2.getRange(4, 1, 1, LLheaders.length).setValues([LLheaders]);

          var leaveArray = leaveData.map(function(Lrow){
                return LLheaders.map(function(Lheader){
                      return Lrow[Lheader];
                });              
          });

        sheet2.getRange(5, 1, leaveArray.length, LLheaders.length).setValues(leaveArray);
    ////////////////////

    // 移動創建的文件到指定文件夾
    var file = DriveApp.getFileById(ss.getId());
    file.moveTo(folder);

    //寫入資料-標題
    var headers = Object.keys(data[0]);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    //寫入資料-數據(神之方法, Javascript 的 map 方法), 經過至少10個版本測試, 用這個方法程式碼最少, 看起來較乾淨
     var dataArray = data.map(function(row){
           return headers.map(function(header){
               return row[header];
           });     
     });
  
    //寫入資料-數據
    sheet.getRange(2, 1, dataArray.length, headers.length).setValues(dataArray);


    ////複寫(更新人事資料每月更新母檔)
     var ccc = "母檔ID位置";
     var mss = SpreadsheetApp.openById(ccc);
     var msheet = mss.getSheets()[0];
        mss.rename(fileName);
     var msheet2 = mss.getSheetByName('請假單');
     
     msheet.clear();
     msheet.setName(last4Char);

     var source1 = SpreadsheetApp.openById(ssID).getSheets()[0];
     var soruce1LastRow = source1.getLastRow();
     var sorce1range = source1.getRange(1, 1, soruce1LastRow, headers.length);
     var source1Values = sorce1range.getValues();

     msheet.getRange(1, 1, source1Values.length, source1Values[0].length).setValues(source1Values);

     msheet2.clear();
     var source2 = SpreadsheetApp.openById(ssID).getSheetByName('請假單');
     var source2LastRow = source2.getLastRow();
     var source2range = source2.getRange(1, 1, source2LastRow, LLheaders.length);
     var source2Values = source2range.getValues();

     msheet2.getRange(1, 1, source2Values.length, source2Values[0].length).setValues(source2Values);

     
      ////修正JASON格式轉入, D, E, F column data 後面很多空白的問題
        removeTrailingSpaces();
    return ContentService.createTextOutput('Data Uploaded Successfully');
}
