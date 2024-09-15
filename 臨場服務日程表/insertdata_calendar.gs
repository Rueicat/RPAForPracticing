/*
1.一樣因為google有日曆服務的關係, 安排好下個月行程後, 可以自動帶入到日曆中不用自己key in
2.資料都是從試算表抓取, 每次開啟時會自動檢查(onOpen)
*/


function autocalendar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('第一區醫師費用');
  var cellValues = ss.getRange('A3:I21').getValues();
  var calendar = CalendarApp.getCalendarById('公用日曆帳號@gmail.com');



  for (var i = 1; i < cellValues.length; i += 2) {
    var currentRow = cellValues[i];
    var nextRow = cellValues[i + 1];
    var date = new Date(currentRow[0]);
    var regionCode = currentRow[7];
    var shop = "(" + currentRow[2] + "、" + nextRow[2] + ")";
    var who = currentRow[8];
    var eventTitle = getEventTitle(regionCode, shop);

    if (eventTitle && !eventExists(calendar, date, eventTitle)) {
      var event = calendar.createAllDayEvent(eventTitle, date);
      event.removeAllReminders(); // 移除該事件的所有提醒
      var eventWho = calendar.createAllDayEvent(who, date);
      eventWho.removeAllReminders();



    }
  }
}

function getEventTitle(regionCode, shop) {
  switch (regionCode) {
    case " ":
      return "one_醫師臨場" + shop;
    case "two醫師執行":
      return "two_醫師臨場" + shop;
    case "three醫師執行":
      return "three_醫師臨場" + shop;
    default:
      return null;
  }
}

function eventExists(calendar, date, title) {
  var events = calendar.getEventsForDay(date);
  for (var i = 0; i < events.length; i++) {
    if (events[i].getTitle() === title) {
      return true;
    }
  }
  return false;
}
