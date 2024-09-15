/////2024.4.11 加入圖片顯示, 手機可以, 電腦如果上傳兩張圖片以上無法顯示...之後再研究
/////2024.5.31 後來發現都會出現!
/////2024.6.21 修改所屬單位取值左邊四位, 因為有的clients還是會填入中文字, 這樣修改可以避免錯誤顯示
////因為表格的設計, 可以分開一般傷病通報和事故通報. 

///////////////////////全域宣告
var dsfolderID = 'destinationID';  // 資料處理處
var dsfolder = DriveApp.getFolderById(dsfolderID);
var ss = SpreadsheetApp.getActiveSpreadsheet();
var jpgID;
var 要寄信的pdf;
var 圖片的ID;


var 處理好的文件ID;
var sheet = ss.getSheetByName('表單回應 1');
var 個資sheet = ss.getSheetByName('(勿動)人事資料表');
var 個資lastrow = 個資sheet.getLastRow();
var lastrow = sheet.getLastRow();
var 時間戳記 = sheet.getRange(lastrow, 1).getValue();
var 時間戳記改 = Utilities.formatDate(時間戳記, "GMT+08:00", "yyyy/MM/dd");
     
 
var 員編 = sheet.getRange(lastrow, 2).getValue();
var 員編搜尋 = 員編;
var 姓名 = sheet.getRange(lastrow, 3).getValue();
var 身分證號碼 = sheet.getRange(lastrow, 4).getValue();
var 所屬單位全 = sheet.getRange(lastrow, 5).getValue();
var 所屬單位 = 所屬單位全.substring(0, 4);
var 所屬單位搜尋 = 所屬單位;
var 通報原因 = sheet.getRange(lastrow, 6).getValue();   /////一般傷病和事故選擇判斷

///一般傷病區
var 事發時間傷 = sheet.getRange(lastrow, 7).getValue();
var 事發時間傷改;
    if(事發時間傷 instanceof Date){
      事發時間傷改 = Utilities.formatDate(事發時間傷, "GMT+08:00", "yyyy/MM/dd");
    }else{事發時間傷改 = ""};
 
  
 
var 事件經過說明傷 = sheet.getRange(lastrow, 8).getValue();
var 受傷部位及程度說明傷 = sheet.getRange(lastrow, 9).getValue();
var 夥伴現況傷 = sheet.getRange(lastrow, 10).getValue();
var 醫院名稱傷 = sheet.getRange(lastrow, 11).getValue();

/////事故
var 事發時間工 = sheet.getRange(lastrow, 12).getValue();
var 事發時間工改;
   if (事發時間工 instanceof Date){
      事發時間工改 = Utilities.formatDate(事發時間工, "GMT+08:00", "yyyy/MM/dd");
   }else{事發時間工改 = ""};

var 災害類型工 = sheet.getRange(lastrow, 13).getValue();
var 車號工 = sheet.getRange(lastrow, 14).getValue();
var 事發地點 = sheet.getRange(lastrow, 15).getValue();
var 事件經過說明工 = sheet.getRange(lastrow, 16).getValue();
var 受傷部位工 = sheet.getRange(lastrow, 17).getValue();
var 相關設備工 = sheet.getRange(lastrow, 18).getValue();
var 夥伴現況工 = sheet.getRange(lastrow, 19).getValue();
var 醫院名稱工 = sheet.getRange(lastrow, 20).getValue();
var 不安全行為工 = sheet.getRange(lastrow, 21).getValue();
var 預防改善工 = sheet.getRange(lastrow, 22).getValue();
var 預計改善完成工改;
   if(預防改善工 instanceof Date){
     預計改善完成工改 = Utilities.formatDate(預計改善完成工, "GMT+08:00", "yyy/MM/dd");
        }else{預計改善完成工改 = ""}

var 填表主管工 = sheet.getRange(lastrow, 24).getValue();
var 附件照片 = sheet.getRange(lastrow, 25).getValue();

///個資搜尋 vlookup with 員編
const vdata = 個資sheet.getRange(1, 1, 個資lastrow, 16).getValues();
let found = false;///chapgpt的方法 
    for(n = 0; n < vdata.length; n++){
        if(vdata[n][0]==員編搜尋){
          found = true;  ///if 找到匹配, change found to TRUE
          break;
        }
    }
    /////According to the found, setting it's value
var 性別設定 = found ? vdata[n][10] : "";            
var 年齡設定 = found ? vdata[n][15] : "";                 

var 性別 = 性別設定;
var 年齡 = 年齡設定;

//////////////////////額外跑單位名稱
const dian = ss.getSheetByName('(勿動)部區店鋪表');
const lastdian = dian.getLastRow();
const vdian = dian.getRange(1, 1, lastdian, 2).getValues();
//////2.20修改
 let founds = false;
    for (v = 0; v < vdian.length; v++){
      if(vdian[v][0]==所屬單位搜尋){
        founds = true;   ///if find data, set it to true
        break;
        }
    }
var 單位名稱設定 = founds ? vdian[v][1] : "總部夥伴";

var 單位名稱 = 單位名稱設定;


async function autoreply() {
////先刪除資料夾的檔案
  const oldfiles = DriveApp.getFolderById(dsfolderID).getFiles();
    while(oldfiles.hasNext()){
      oldfiles.next().setTrashed(true);
    }
Utilities.sleep(5000);
////複製母檔到新資料夾處理
const originaldoc = DriveApp.getFileById('1KvKwrblXnlSOr15_PXRFd2SzZyxEUAc6s_SRUcL-R90');   ///母檔DOC ID
var processdocID = originaldoc.makeCopy('處理_事件通報', dsfolder).getId();    ///複製出要編輯的DOC
var processdoc = DocumentApp.openById(processdocID);
var olddocbody = processdoc.getBody();

olddocbody.replaceText("{{時間戳記}}", 時間戳記改);
olddocbody.replaceText("{{員編}}", 員編);
olddocbody.replaceText("{{姓名}}", 姓名);
olddocbody.replaceText("{{性別}}", 性別);
olddocbody.replaceText("{{年齡}}", 年齡);
olddocbody.replaceText("{{所屬單位}}", 所屬單位);
olddocbody.replaceText("{{所屬單位名稱}}", 單位名稱);
olddocbody.replaceText("{{受傷部位及受傷程度說明}}", 受傷部位工 + 受傷部位及程度說明傷);
olddocbody.replaceText("{{夥伴現況}}", 夥伴現況工 + 夥伴現況傷);
olddocbody.replaceText("{{醫院名稱}}", 醫院名稱工 + 醫院名稱傷);
olddocbody.replaceText("{{通報原因}}", 通報原因);
olddocbody.replaceText("{{事發時間}}", 事發時間工改 + 事發時間傷改);
olddocbody.replaceText("{{事發地點}}", 事發地點);
olddocbody.replaceText("{{災害類型}}", 災害類型工);
olddocbody.replaceText("{{相關設備}}", 相關設備工);
olddocbody.replaceText("{{請填寫交通事故「車號」}}", 車號工);
olddocbody.replaceText("{{事件經過說明}}", 事件經過說明工 + 事件經過說明傷);
olddocbody.replaceText("{{夥伴或他人是否有下列「不安全行為或環境」造成事故}}", 不安全行為工);
olddocbody.replaceText("{{預防及改善措施}}", 預防改善工);
olddocbody.replaceText("{{預防及改善措施預計完成日}}", 預計改善完成工改);
olddocbody.replaceText("{{填表主管}}", 填表主管工);
olddocbody.replaceText("{{受傷照片或診斷書上傳}}", 附件照片);

processdoc.saveAndClose();

處理好的文件ID = processdocID;

///轉成PDF
const DOCtopdfBLOB = DriveApp.getFileById(processdocID).getAs('application/pdf');
要寄信的pdf = DriveApp.getFolderById(dsfolderID).createFile(DOCtopdfBLOB).setName(員編 + "_" + 姓名 + "_" + 所屬單位 + "_" + 單位名稱);

await main_pdf_to_pic();    //////等待確定產出圖片, 以下抓取圖片ID

////////////////////////////////////////////////////////以下確認google drive 有更新列表
const timeout = 20000; // 超過的時間--20秒
  const startTime = new Date().getTime();
  let found = false;

  while (new Date().getTime() - startTime < timeout) {           ////////////經過時間 - 開始時間  <  設置時間
    const query = dsfolder.getFilesByName('page1.png');
    if (query.hasNext()) {
      圖片的ID = query.next().getId();
      
      found = true;
      break;
    }
    // 等待一段時間再重试
    Utilities.sleep(2000); // 每次重試間隔2秒
  }

  if (!found) {
    
    SpreadsheetApp.getUi().alert("未在指定時間找到圖片檔, 應該是伺服器忙線, 請重新執行");
    return;
  }


await autoemail();
}



