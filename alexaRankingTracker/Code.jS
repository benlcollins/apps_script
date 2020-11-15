/**
 * save new row of alexa data
 */
function saveAlexaData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("alexa_rankings");
  const settingsSheet = ss.getSheetByName("settings");
  
  // get the url, follower count and date from the first three cells
  const d = new Date();
  const datestring = (d.getMonth()+1) + "/" + d.getDate()  + "/" + d.getFullYear();
  const global_count = settingsSheet.getRange(11,2).getValue();
  const us_count = settingsSheet.getRange(11,3).getValue();
  
  // append new ranking data to Sheet
  dataSheet.appendRow([
    datestring,
    global_count,
    us_count
  ]);
}