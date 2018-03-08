function onOpen() {
  
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Index Menu')
      .addItem('Create Index', 'createIndex')
      .addItem('Update Index', 'updateIndex')
      .addToUi();
}


// function to create the index
function createIndex() {
  
  // Get all the different sheet IDs
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  
  var namesArray = sheetNamesIds(sheets);
  
  var indexSheetNames = namesArray[0];
  var indexSheetIds = namesArray[1];
  
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
    
    printIndex(indexSheet,indexSheetNames,indexSheetIds);

  }
    
}



// function to update the index, assumes index is the first sheet in the workbook
function updateIndex() {
  
  // Get all the different sheet IDs
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var indexSheet = sheets[0];
  
  var namesArray = sheetNamesIds(sheets);
  
  var indexSheetNames = namesArray[0];
  var indexSheetIds = namesArray[1];
  
  printIndex(indexSheet,indexSheetNames,indexSheetIds);
}


// function to print out the index
function printIndex(sheet,names,formulas) {
  
  sheet.clearContents();
  
  sheet.getRange(1,1).setValue('Workbook Index').setFontWeight('bold');
  sheet.getRange(3,1,names.length,1).setValues(names);
  sheet.getRange(3,2,formulas.length,1).setFormulas(formulas);
  
}


// function to create array of sheet names and sheet ids
function sheetNamesIds(sheets) {
  
  var indexSheetNames = [];
  var indexSheetIds = [];
  
  // create array of sheet names and sheet gids
  sheets.forEach(function(sheet){
    indexSheetNames.push([sheet.getSheetName()]);
    indexSheetIds.push(['=hyperlink("#gid=' 
                        + sheet.getSheetId() 
                        + '","' 
                        + sheet.getSheetName() 
                        + '")']);
  });
  
  return [indexSheetNames, indexSheetIds];
  
}
