/**
 * menu
 */
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('Sidebar Menu')
      .addItem('Show sidebar', 'showSidebar')
      .addToUi();
}

/**
 * show sidebar
 */
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('sidebar.html')
      .evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Index Sidebar');
  
  SpreadsheetApp.getUi().showSidebar(ui);
}


/**
 * get the sheet names
 */
function getSheetNames() {
  
  // Get all the different sheet IDs
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  
  return sheetNamesIds(sheets);
}


// function to create array of sheet names and sheet ids
function sheetNamesIds(sheets) {
  
  var indexOfSheets = [];
  
  // create array of sheet names and sheet gids
  sheets.forEach(function(sheet){
    indexOfSheets.push([sheet.getSheetName(),sheet.getSheetId()]);
  });
  
  //Logger.log(indexOfSheets);
  return indexOfSheets;
  
}