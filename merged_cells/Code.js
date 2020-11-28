/**
 * Merged Cells Utility Script
 * Helps user identify all merged cells in a Google Sheet
 * 
 * To Do:
 * Return background color of merged cells to original color, if not white
 */

// menu to run merged Cell highlighter
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Merged Cells Finder')
      .addItem('Highlight merged cells', 'highlightMergedCells')
      .addItem('Un-highlight merged cells', 'unhighlightMergedCells')
      .addToUi();
}

// function to find all sheets to highlight merged cells
function highlightMergedCells() {
  
  // get array of Sheets
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // variable to hold merged cells info
  let mergedCellInfo = '';

  // find merged cells in each
  sheets.forEach(sheet => {

    // find merged cell info for this sheet
    const newMergedCellInfo = highlightMergedCellsInSheet(sheet);
    
    // add new merged cell info to existing info
    mergedCellInfo = mergedCellInfo + newMergedCellInfo;

  });

  // output summary pane of merged cells info
  SpreadsheetApp.getUi().alert(mergedCellInfo);
}

// function to find all sheets to un-highlight merged cells
function unhighlightMergedCells() {
  
  // get array of Sheets
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // loop over all sheets in array
  sheets.forEach(sheet => {

    // pass sheet to function to unhighlight merged cells
    unhighlightMergedCellsInSheet(sheet);

  });

}

// highlight merged cells in specific sheet
function highlightMergedCellsInSheet(sheet) {
  
  // get the merged ranges
  const range = sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns());
  const mergedRanges = range.getMergedRanges();

  // string variable to hold merged cell info
  let mergedCellString = '';
  
  // loop over merged ranges to highlight and return details
  mergedRanges.forEach((rng,i) => {

    // highlight merged cells with yellow background
    rng.setBackground('yellow');

    // gather the merged cells info and add to string
    mergedCellString = mergedCellString + 'Merged Cell Group ' + sheet.getName() + ' (' + (i+1) + '):\n'; // which merged cell group
    mergedCellString = mergedCellString + 'Range: ' + sheet.getName() + '!' + rng.getA1Notation() + '\n'; // what is the merged cells range
    mergedCellString = mergedCellString + 'Contents: ' + rng.getDisplayValue() + '\n\n'; // include the merged cells contents
  
  });

  // return merged cell info to main function
  return mergedCellString;

}

// un-highlight merged cells in sheet
function unhighlightMergedCellsInSheet(sheet) {

  // get the merged ranges
  const range = sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns());
  const mergedRanges = range.getMergedRanges();
  
  // loop over merged ranges in Sheet
  mergedRanges.forEach(rng => {

    // set background to white
    rng.setBackground('white');

  });

}