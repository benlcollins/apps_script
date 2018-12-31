// add custom menu to run from Sheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Calendar App')
    .addItem('Upload to Calendar','sheetsToCalendar')
    .addToUi();
}

// function to retrieve data from Sheet and add to Calendar
function sheetsToCalendar() {
  
  // get spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Fixtures');

  // ask user which calendar they want to add data to
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
      'What calendar do you want to upload these dates to?',
      'Please enter your email:',
      ui.ButtonSet.OK_CANCEL);

  var email = result.getResponseText();

  // get calendar
  var masterCal = CalendarApp.getCalendarById(email);
  
  var fixtures = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  
  fixtures.forEach(function(fixture) {
    
    var startTime = fixture[0];
    var endTime = new Date(startTime);
    addHours(endTime,2);
    var title = fixture[1] + ' - ' + fixture[2] + ' v ' + fixture[3];

    Logger.log(startTime);
    Logger.log(endTime);
    Logger.log(title);

    try {
      masterCal.createEvent(title,startTime,endTime);  
    } catch(e) {
      Logger.log('Error with calendar (' + email + '): ' + e);
    }
    
  })
  
}

// add hours to date
function addHours(date,hours) {

  return date.setHours(date.getHours()+ hours);
  
}