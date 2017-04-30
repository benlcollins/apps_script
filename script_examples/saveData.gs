// custom menu function
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
      .addItem('Save Data','saveData')
      .addToUi();
}
 
// function to save data
function saveData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var data = sheet.getRange('Sheet1!A2:D2').getValues();
  Logger.log(data);
  sheet.appendRow(data[0]);
}

// function to copy file 
function copySheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.copy("Copy of " + ss.getName());
}