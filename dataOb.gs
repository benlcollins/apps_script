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
  
  // get all the data
  var allData = newSheet.getDataRange().getValues();
  
  // get the header row
  var headers = allData.shift();
  Logger.log(headers);
  
  var headerIndexes = indexifyHeaders(headers);
  Logger.log(headerIndexes);
  
}




function indexifyHeaders(headers) {
  
  var index = 0;
  return headers.reduce (function (p,c) {
    
    // skip column headings with blank headers
    if (c) {
      // throw error if duplicate header names
      if (p.hasOwnProperty (c)) {
        throw new Error('Duplicate column name ' + c);
      }
      p[c] = index;
    }
    index++;
    return p;
  },{});
}
  
  
  
  
  
  
  
