/**
* Custom menu
*/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sheet Size Auditor')
      .addItem('Audit Sheet', 'sheetAuditor')
      .addToUi();
}

/**
* Audits the given url and displays results into browser message box
*/
function sheetAuditor() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
  
    var audit_data = auditUrl(ss);
    Logger.log(audit_data);
    Browser.msgBox("Number of sheets: " + audit_data[0] + 
                    "\\nTotal cells used: " + audit_data[1] + 
                      "\\nPercent cells used: " + Math.round(audit_data[2]*100) + "%");

  }
  catch(e) {
    Logger.log(e);
  }
}


/**
* Returns size data for a given sheet url
* @param {string} url - url of the Google Sheet
* @returns {array} of file data and individual sheets data
*/
function auditUrl(spreadsheet) {
  
  var sheets = spreadsheet.getSheets();
  
  // counters
  var numSheets = 0;
  var totalCellCounter = 0;
  var fileArray = [];
  
  sheets.forEach(function(sheet) {

    // get single sheet data
    var thisSheetInfo = getSingleSheetInfo(sheet);
    
    // update the counters
    numSheets++;
    totalCellCounter = totalCellCounter + thisSheetInfo[3];
    
  });

  var totalCellPercent = totalCellCounter / 5000000;
  
  fileArray.push(numSheets, totalCellCounter, totalCellPercent);
  
  return fileArray;
 
}


/**
* Returns individual sheet size data
* @param {object} sheet - the individual sheet within a Google Sheet
* @returns {array} of configuration settings
*/
function getSingleSheetInfo(sheet) {
    
  var singleSheetArray = [];

  var name = sheet.getName();
   
  // how many cells in the sheet currently
  var maxRows = sheet.getMaxRows();
  var maxCols = sheet.getMaxColumns();
  
  var totalCells = maxRows * maxCols;
    
  singleSheetArray.push(
    name,
    maxRows,
    maxCols,
    totalCells
  );
  
  return singleSheetArray;
}
