# 人事資料處理自動模組

workflow: 

* 選擇local excel 檔案後, 會自動上傳到指定google drive
* 上傳前會編輯內容, 改成需要的內容(例如: 增加年齡, 年資等)
* 更新線上人事資料
  * 一份編碼成spreadsheet線上母檔, 複寫過去並更新sheet name至更新日期
  * 另一份會原始excel檔案備份到google drive指定地點 
* 原始人事資料sheet[0]預設是人事資料, sheet[1]是請假資料, 這邊透過遠端apps script做更新(因為就全部copy-past 就好, 練習不同寫法)