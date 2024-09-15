function datadata() {
       
   var ddd = Browser.msgBox('確定幫你填寫日期時間?', Browser.Buttons.OK_CANCEL);
   if (ddd === 'ok'){



   var ssgo = ss.getSheetByName('臨場服務明細表');
  ///XXXX鄰廠服務日程表
  var ssshop = SpreadsheetApp.openById('ur ID');
  var arrangeshop = ssshop.getSheetByName('XX醫師費用');

  ////店鋪代號及臨場日期
  var datarange = arrangeshop.getRange(4, 1, 18, 2).getValues();
              
  ///Target shopcode Q2
  var target = ssgo.getRange(2, 17, 1, 1).getValue();
         
  
  //////for circle
      for (var i = 0; i < datarange.length; i ++){
             if(datarange[i][1]==target){break}
      }  
                     
          
 
       ////判斷第一家店還是第二家店, 並擷取日期
         if(i % 2 == 0){
            ssgo.getRange(3, 5, 1, 1).setValue('13時');

            var 臨場日期 = datarange[i][0];
             var monthnumber = 臨場日期.getMonth() + 1;                
            ssgo.getRange(3, 3, 1, 1).setValue(monthnumber + '月');
             var daynumber = 臨場日期.getDate();
            ssgo.getRange(3, 4, 1, 1).setValue(daynumber + '日');

         } else {
             ssgo.getRange(3, 5, 1, 1).setValue('15時');
               
              var 臨場服務第二家 = datarange[i -1][0]; 
               var monthnumber = 臨場服務第二家.getMonth() + 1;    
             ssgo.getRange(3, 3, 1, 1).setValue(monthnumber + '月');
                var daynumber = 臨場服務第二家.getDate();
             ssgo.getRange(3, 4, 1, 1).setValue(daynumber + '日');
         }      
    

   

    ss.toast('日期時間確認好嚕~', '💡提醒小視窗', 5)

   }
}

