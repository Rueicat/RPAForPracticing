# 工傷通報事件紀錄流程自動化
能依照spreadsheet的資料, 一鍵寄出信件給各單位相關人員, 寄出後會在checkbox 打勾紀錄
* 自動依照員工所屬單位, 找查所有該員工所屬需通報單位
* 自動創建表格並自動填寫通報內容, 於信件中發送
* 簽名檔會依照登入身分切換(需事先設定與設計簽名檔)
  > note: 需要使用google 服務(這是用google apps script設計出來的客製化RPA)

## UI介面設計
* 使用內建程式庫, 直接在spreadsheet 上方選單增加客製化選項, 如下圖所示:

![custom_list](./工傷通報事件紀錄流程自動化/custom_list.png)

* 信件自動表單如下所示:

  內容都是自動從spreadsheet抓取, 詳見`Main code.gs` 內容

![form_preview](./工傷通報事件紀錄流程自動化/form_preview.png)

* 簽名檔設計是基本的html語法

  (略)

* spreadsheet 項目參考如下:

  有使用到的功能:
  
  * iferror + vlookup 的組合
  * google spreadsheet 的下拉選單
  * 內建的checkbox功能(可以判斷true or false)

![spreadsheetForm_exammple](./工傷通報事件紀錄流程自動化/spreadsheetForm_example.png)

