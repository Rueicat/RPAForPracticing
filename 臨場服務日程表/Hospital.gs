/*
1.google服務中有map, 所以可以簡單利用apps scirpt使用這服務計算里程數
2.路徑是依照google判定的最佳路徑
3.當名稱很類似目的地輸入時, map 可能會判定錯誤, 如果能提供地址比較安全
4.這是使用onOpen, 所以打開表單會自動更新(會這樣選擇是因為排形成通常一個月一次),避免再設計一個"計算里程數的按鈕"讓使用者啟動(希望讓user動最少的手指)
*/


function Map_distance() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('醫師費用');
  var VGH = ss.getRange('B7').getValue();
  var firstshop ="名稱 " + ss.getRange('B4').getValue();
  var secondshop ="名稱 " + ss.getRange('B5').getValue();
  
  ////醫院出發到第一家店
 var VGHtofirstshop = Maps.newDirectionFinder()
                      .setOrigin(VGH).setDestination(firstshop)
                      .setMode(Maps.DirectionFinder.Mode.DRIVING)
                      .getDirections();
var Distance_of_VGH_to_firstshop = VGHtofirstshop.routes[0].legs[0].distance.text;   /////routes[0]是指google建議的路線中取第一個建議的路線(最近的),legs是取第一段路途(如果中間有很多中途地點)
var only_Distance_of_VGH_to_firstshop = Distance_of_VGH_to_firstshop.match(/[\d\.]+/)[0];   ///他會顯示km, 但前端試算表其他格子的公式只能抓數字, 所以利用regular expresion 抓數字部分
ss.getRange('E7').setValue(only_Distance_of_VGH_to_firstshop); 


///第一家店到第二家
 var firsttosecondshop = Maps.newDirectionFinder()
                      .setOrigin(firstshop).setDestination(secondshop)
                      .setMode(Maps.DirectionFinder.Mode.DRIVING)
                      .getDirections();

var Distance_of_first_to_secondshop = firsttosecondshop.routes[0].legs[0].distance.text;
var only_Distance_of_first_to_secondshop = Distance_of_first_to_secondshop.match(/[\d\.]+/)[0];  
ss.getRange('E8').setValue(only_Distance_of_first_to_secondshop); 



///第二家店回醫院
var backtoVGH = Maps.newDirectionFinder()
                      .setOrigin(secondshop).setDestination(VGH)
                      .setMode(Maps.DirectionFinder.Mode.DRIVING)
                      .getDirections();

var Distance_of_backtoVGH = backtoVGH.routes[0].legs[0].distance.text;
var only_Distance_of_backtoVGH = Distance_of_backtoVGH.match(/[\d\.]+/)[0];  
ss.getRange('E9').setValue(only_Distance_of_backtoVGH); 



}
