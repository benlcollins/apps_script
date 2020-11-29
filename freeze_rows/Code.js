/**
 * Frozen Rows and Columns Helper Functions
 * Ben Collins
 * November 2020
 */

// custom menu to operate frozen rows helper from Sheet
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Freeze Rows Helper')
      .addItem('Identify Freeze Rows or Columns', 'identifyAllFrozenRowsColumns')
      .addItem('Remove All Freeze Rows or Columns', 'removeAllFrozenRowsColumns')
      .addToUi();
}


function identifyAllFrozenRowsColumns() {
  
  // get array of Sheets
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // variable to hold freeze rows info
  let freezeInfo = '';

  // find merged cells in each
  sheets.forEach(sheet => {

    // find merged cell info for this sheet
    const newFreezeInfo = findFrozenRowsColumnsInSheet(sheet);
    
    // add new merged cell info to existing info
    freezeInfo = freezeInfo + newFreezeInfo;

  });

  // output summary pane of merged cells info
  SpreadsheetApp.getUi().alert(freezeInfo);
}

// find frozen rows and columns in specific sheet
function findFrozenRowsColumnsInSheet(sheet) {
  
  // get the merged ranges
  const sheetName = sheet.getName();
  const frozenRows = sheet.getFrozenRows();
  const frozenCols = sheet.getFrozenColumns();

  // string variable to hold merged cell info
  let freezeString = '';

  freezeString = freezeString + 'Sheet: ' + sheetName + '\n'; // which Sheet
  freezeString = freezeString + 'Frozen rows: ' + frozenRows + '\n'; // how many frozen rows
  freezeString = freezeString + 'Frozen columns: ' + frozenCols + '\n\n'; // how many frozen columns

  // return merged cell info to main function
  return freezeString;

}

function removeAllFrozenRowsColumns() {

  // get array of Sheets
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // find merged cells in each
  sheets.forEach(sheet => {

    sheet.setFrozenRows(0);
    sheet.setFrozenColumns(0);

  });

}
