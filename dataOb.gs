// abstracting header row in javascript object
// avoid referencing columns by number
function dataOb() {
  
  // get this workbook
  var thisWorkbook = SpreadsheetApp.getActiveSpreadsheet();
  
  // get this sheet
  var thisSheet = thisWorkbook.getSheetByName('carriers');
  
  // check if sheet exists and delete it if it does
  var newSheet = thisWorkbook.getSheetByName('dataObWorkingSheet');
  if (newSheet) {
    thisWorkbook.deleteSheet(newSheet);
  }
  
  // make a copy of the sheet
  var newSheet = thisSheet.copyTo(thisWorkbook);
  
  // set a new name
  newSheet.setName('dataObWorkingSheet');
  
  // make it the active sheet
  thisWorkbook.setActiveSheet(newSheet);
  
  
  
}
