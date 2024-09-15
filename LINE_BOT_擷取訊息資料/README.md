# LINE BOT 擷取訊息

公司常會利用line通報事件, 這可以利用LINE BOT 自動擷取訊息, 貼到指定的表單內查閱紀錄

## line 官方文件有說明相關內容, 以下自己的memo...

* `Event.Type`

* `Event.message.type`
  * Text (感覺文字訊息比較常會用到)
  * Image
  * Video
  * Audio
  * File
  * Location
  * Sticker

* Message event outer shell
  * type: message
  * reply Token: only for 30mins and can use it once
  * message: object

> 接收資料是JSON格式