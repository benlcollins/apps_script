// change this to be whatever value you want
var THRESHOLD_VALUE = 100;

// function to check whether value is greater than threshold
// set to run whenever sheet is edited, or at short set time intervals
// whichever works better
function valueNotification() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var valueSheet = ss.getSheetByName('Sheet1');
  var emailSheet = ss.getSheetByName('Emails');
  
  // get current value from sheet
  var value = valueSheet.getRange(1,1).getValue();
  
  // get list of emails from sheet
  var emails = emailSheet.getRange(1,1,emailSheet.getLastRow(),1).getValues();
  var emailString = [].concat.apply([], emails).join();
  
  if (value > THRESHOLD_VALUE) {
    sendEmail(value,emailString);
  }
}


function sendEmail(value,emails) {
  
  // get current date
  var now = new Date();
  
  // send emails to recipients listed in sheet
  GmailApp.sendEmail(emails, "Value changed in Value notification Sheet: " + now, "New value is: "+
                     value + ", which is above the threshold value of: " + THRESHOLD_VALUE)
  
}