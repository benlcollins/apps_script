function sheetsToCalendar() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var masterCal = CalendarApp.getCalendarById('ben@benlcollins.com');
  
  var fixtures = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  
  //Logger.log(fixtures);
  
  fixtures.forEach(function(fixture) {
    
    var startTime = fixture[0];
    var endTime = new Date(startTime);
    addHours(endTime,2);

    Logger.log(startTime);
    Logger.log(endTime);
    Logger.log(fixture[1] + ' - ' + fixture[2] + ' v ' + fixture[3]);
    
    
  })
  
}


function addHours(date,hours) {

  return date.setHours(date.getHours()+ hours);
  
}