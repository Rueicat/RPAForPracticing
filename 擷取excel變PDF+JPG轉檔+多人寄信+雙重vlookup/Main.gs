////////全域宣告區
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('通知單');
var 月 = sheet.getRange(5, 8, 1, 2).getValues()[0][0];
var 日 = sheet.getRange(5, 8, 1, 2).getValues()[0][1];
var 店鋪代號 = sheet.getRange(4, 3, 1, 2).getValues()[0][0];
var 寄信店鋪信箱 = 店鋪代號.replace("B", "mos") + "@mos.com.tw";
var 店鋪名稱 = sheet.getRange(4, 3, 1, 2).getValues()[0][1];
var dsfolder = DriveApp.getFolderById('destinationFolderID'); ///程式作業區
var 通知單ID;
var 圖片的ID;
var 要寄信的pdf;
//////

//////////////ok要小寫, 因為blob無法單獨打包指定sheet(需要整個檔案), 所以先隱藏不需要sheet, 轉檔後再展開
async function sendinform() {
  const confirm = Browser.msgBox('記得刪除匿名圈圈再寄送', Browser.Buttons.OK_CANCEL);
    if(confirm=='ok'){
      const oldfiles = dsfolder.getFiles();
       while(oldfiles.hasNext()){
          oldfiles.next().setTrashed(true);
       }

Utilities.sleep(5000); 

    const allsheet = ss.getSheets();
      for(let i = 0; i < allsheet.length; i++){
        if(allsheet[i].getSheetName() !=='通知單'){    ////驚嘆號指排除這個選項
          allsheet[i].hideSheet()
        }
      }
const pdfBlob = ss.getBlob();
const 通知單blob = dsfolder.createFile(pdfBlob);
      通知單blob.setName(店鋪代號 + '_' + 店鋪名稱 + '_臨場服務通知單_' + 月 + 日);
要寄信的pdf = 通知單blob;

通知單ID = 通知單blob.getId();
      for(let i = 0; i < allsheet.length; i++){
        allsheet[i].showSheet();
      }
     
await main_pdf_to_pic();    ///等待確定產出圖片, 以下抓取圖片ID

////////////////////////////////////////////////////////以下確認google drive 有更新列表
const timeout = 20000; // 超過幾秒就算失敗，20秒
  const startTime = new Date().getTime();
  let found = false;

  while (new Date().getTime() - startTime < timeout) {           ////////////經過時間 - 開始時間  <  設置時間
    const query = dsfolder.getFilesByName(店鋪代號 + '_' + 店鋪名稱 + '-臨場服務通知單-' + 月 + 日);
    if (query.hasNext()) {
      圖片的ID = query.next().getId();
      
      found = true;
      break;
    }
    // 等待伺服器回應, 再重試
    Utilities.sleep(2000); // recheck sleep 2 interval times
  }

  if (!found) {
    
    SpreadsheetApp.getUi().alert("未在指定時間找到圖片檔, 應該是伺服器忙線, 請重新執行");
    return;
  }

await double_vlookup();
await autoemail();

    ss.toast('~已經寄出了喔~', '💡提醒小視窗', 5);
    }else{}
}
