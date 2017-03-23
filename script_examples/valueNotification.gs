// change this to be whatever value you want
var THRESHOLD_VALUE = 100;

// function to check whether value is greater than threshold
// set to run whenever sheet is edited, or at short set time intervals
// whichever works better
function valueNotification() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var valueSheet = ss.getSheetByName('Sheet1');
  var emailSheet = ss.getSheetByName('Emails');
  
  var value = valueSheet.getRange(1,1).getValue();
  
  var emails = emailSheet.getRange(1,1,emailSheet.getLastRow(),1).getValues()[0];
  Logger.log(value);
  Logger.log(emails);
  
  if (value > THRESHOLD_VALUE) {
    sendEmail(value,emails);
  }
}


function sendEmail(value,emails) {
  
  // get current date
  var now = new Date();
  Logger.log(now);
  Logger.log(value);
  Logger.log(emails);
  
}