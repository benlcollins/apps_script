function onOpen() {
  
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Index Menu')
      .addItem('Create Index', 'createIndex')
      .addToUi();
}

function createIndex() {
  
  // Log all the different sheet IDs
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  
  var indexSheetNames = [];
  var indexSheetIds = [];
  
  // create array of sheet names and sheet gids
  sheets.forEach(function(sheet){
    indexSheetNames.push([sheet.getSheetName()]);
    indexSheetIds.push(['=hyperlink("https://docs.google.com/spreadsheets/d/1D7nX_l3wlNEAoBuEyde5daQgMy56uPMco0aVs51ibVY/edit#gid=' 
                        + sheet.getSheetId() 
                        + '","' 
                        + sheet.getSheetName() 
                        + '")']);
  });
  
  // check if sheet called sheet called already exists
  // if no index sheet exists, create one
  if (ss.getSheetByName('index') == null) {
    
    var indexSheet = ss.insertSheet('Index',0);
    
  }
  // if sheet called index does exist, prompt user for a different name or option to cancel
  else {
    
    var indexNewName = Browser.inputBox('The name Index is already being used, please choose a different name:', 'Please choose another name', Browser.Buttons.OK_CANCEL);
    
    if (indexNewName != 'cancel') {
      var indexSheet = ss.insertSheet(indexNewName,0);
    }
    else {
      Browser.msgBox('No index sheet created');
    }
    
  }
  
  // add sheet title, sheet names and hyperlink formulas
  if (indexSheet) {
    indexSheet.getRange(1,1).setValue('Workbook Index').setFontWeight('bold');
    indexSheet.getRange(3,1,indexSheetNames.length,1).setValues(indexSheetNames);
    indexSheet.getRange(3,2,indexSheetIds.length,1).setFormulas(indexSheetIds);
  }
    
}

