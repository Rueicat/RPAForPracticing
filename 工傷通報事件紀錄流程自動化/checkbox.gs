////this is Javascript method
///沒有納入apps script指令庫, 但是可以通用

function onEdit(e) {   ///onEdit是特殊用法, 編輯的時候就會啟動以下程式碼, e 是接收你編輯的資料訊息
  var erange = e.range;     ///定義 e (你編輯的資料) 抓出範圍  (就你編輯那個格子的資料範圍, 只有範圍)
  var esheet = erange.getSheet();   ////定義  告訴電腦  你正在編輯的sheet 是哪一個
  var editedColumn = erange.getColumn();   ///  你正在編輯的格子 取 他的  欄column 位子
  var editedRow = erange.getRow();        ///   你正在編輯的格子 取 他的  列row 位子
  var newValue = e.value;               ///抓取 你正在編輯格子的資料


   ///文法：  如果 正在編輯的欄位是 第18欄(就是R欄), 而且 正在編輯的格子, 資料是TRUE (核取方塊打勾對電腦來說是TRUE) 

  if (editedColumn === 18 && newValue === "TRUE"){

    ///將該欄位全部的列row設定成灰色 (網路可以查顏色代碼, 這是"六元碼", 就是#+另個代號 , #d3d3d3 是亮灰色, 網路很好查到)

   esheet.getRange(editedRow, 1, 1, esheet.getLastColumn()).setBackground("#d3d3d3");

    ///如果核取方塊是沒有打勾(FALSE), 該橫列全部變白色

  }else if (editedColumn === 18 && newValue === "FALSE"){

   esheet.getRange(editedRow, 1, 1, esheet.getLastColumn()).setBackground("#FFFFFF");
  }
}
