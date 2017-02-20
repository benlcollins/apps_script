// menu
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('Sidebar Menu')
      .addItem('Show sidebar', 'showSidebar')
      .addToUi();
}

// show sidebar
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
  
  Logger.log(indexOfSheets);
  return indexOfSheets;
  
}

/* TO DO:
 * have menu in Google Sheet to show/hide index
 * automatically create new index when show is run, so it's always up-to-date
 * collect all sheet names into an array
 * pass array of sheet names to HTML service sidebar
 * display all sheet names in HTML sidebar
 * convert all sheet names in sidebar to hyperlinks to different sheets
 * have close button in sidebar
 * add basic CSS styling
 */