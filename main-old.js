function main() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('birthdays');
  var lastRow = sheet.getLastRow();

  for (var i = 2; i <= lastRow; i++) {
    var birthLastName = sheet.getRange(i, 1).getValue();
    if (birthLastName == "") {
      break;
    }

    var registed = sheet.getRange(i, 6).getValue();
    if (registed == "") {
      var today = new Date();
      var todayYear = today.getYear();
      var birthMonth = sheet.getRange(i, 4).getValue();
      var birthDate = sheet.getRange(i, 5).getValue();
      var birthDay = todayYear + "/" + birthMonth + "/" + birthDate;
      var birthFirstName = sheet.getRange(i, 2).getValue();
      var birthMember = birthLastName + " " + birthFirstName;
      var eventTitle = birthMember + " さん誕生日";
      var eventDesc = birthMember + "さん、お誕生日おめでとう！ ";


      // カレンダー ID はスクリプトのプロパティに登録した内容を取得
      var calId = PropertiesService.getScriptProperties().getProperty("clendarId");
      var cal = CalendarApp.getCalendarById(calId);

      // カレンダーに登録 (毎年繰り返す)
      eventSeries = cal.createAllDayEventSeries(eventTitle, new Date(birthDay),
        CalendarApp.newRecurrence().addYearlyRule(), {description: eventDesc});
      if(eventSeries){
        sheet.getRange(i, 6).setValue("登録済み");
      }else{
        sheet.getRange(i, 6).setValue("登録失敗");
      }
    } else {
      console.info("登録済みです.");
    }
  }
}
