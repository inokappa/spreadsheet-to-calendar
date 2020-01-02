function getBirthdays() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('birthdays');
  var values = sheet.getDataRange().getValues(); 
  return values;
}

function registToCalendar(i, event) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('birthdays');
  // カレンダー ID はスクリプトのプロパティに登録した内容を取得
  var calId = PropertiesService.getScriptProperties().getProperty("clendarId");
  var cal = CalendarApp.getCalendarById(calId);
  var eventId = event['id'];
  var eventTitle = event['t'];
  var birthDay = event['dt'];
  var eventDesc = event['ds'];

  // カレンダーに登録 (毎年繰り返す)
  eventSeries = cal.createAllDayEventSeries(eventTitle, new Date(birthDay),
    CalendarApp.newRecurrence().addYearlyRule(), {description: eventDesc});
  if(eventSeries){
    sheet.getRange(eventId, 6).setValue("登録済み");
  }else{
    sheet.getRange(eventId, 6).setValue("登録失敗");
  }
}

function main() {
  var birthdays = getBirthdays();
  for (var i = 1; i < birthdays.length; i++) {
    var birthLastName = birthdays[i][0];
    if (birthLastName == "") {
      break;
    }

    var registed = birthdays[i][5];
    if (registed == "") {
      var today = new Date();
      var todayYear = today.getYear();
      var birthMonth = birthdays[i][3];
      var birthDate = birthdays[i][4];
      var birthDay = todayYear + "/" + birthMonth + "/" + birthDate;
      var birthFirstName = birthdays[i][1];
      var birthMember = birthLastName + " " + birthFirstName;
      var eventTitle = birthMember + " さん誕生日";
      var eventDesc = birthMember + "さん、お誕生日おめでとう！ ";
      var event = { id: i + 1, t: eventTitle, dt: birthDay,  ds: eventDesc };
      registToCalendar(i, event);
    } else {
      console.info("登録済みです.");
    }
  }
}
