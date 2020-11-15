/**
 * save new row of alexa data
 */
function saveAlexaData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("alexa_rankings");
  const settingsSheet = ss.getSheetByName("settings");
  
  // get the url, follower count and date from the first three cells
  const d = new Date();
  const global_count = settingsSheet.getRange(11,2).getValue();
  const us_count = settingsSheet.getRange(11,3).getValue();
  
  // append new ranking data to Sheet
  dataSheet.appendRow([
    d,
    global_count,
    us_count
  ]);

  // format date string cell
  dataSheet.getRange(dataSheet.getLastRow(),1).setNumberFormat("MMM-YY");
}