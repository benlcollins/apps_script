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
  
  // get previous value from sheet
  var previousValue = getPreviousValue() ? getPreviousValue() : 0;
  Logger.log(previousValue);
  
  // get list of emails from sheet
  var emails = emailSheet.getRange(1,1,emailSheet.getLastRow(),1).getValues();
  var emailString = [].concat.apply([], emails).join();
  
  if (value > THRESHOLD_VALUE && value != previousValue) {
    sendEmail(value,previousValue,emailString);
  }
  
  // save value for next time
  storePreviousValue(value);
}


// function to send email
function sendEmail(value,previousValue,emails) {
  
  // get current date
  var now = new Date();
  
  // send emails to recipients listed in sheet
  GmailApp.sendEmail(emails, "Value changed in Value notification Sheet: " + now, "New value is: "+
                     value + ", which is different from the previous value: " + previousValue + " and is above the threshold value of: " + THRESHOLD_VALUE)
  
}


/*
 * script properties service
 *
 */

// retrive copy of previous value
function getPreviousValue() {  
  var properties = PropertiesService.getScriptProperties();
  return properties.getProperty('valueNotificationPreviousValue');
}

// save copy of previous notification value
function storePreviousValue(val) {
  var properties = PropertiesService.getScriptProperties();
  properties.setProperty('valueNotificationPreviousValue', val);
}