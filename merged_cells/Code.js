/**
 * Merged Cells Utility Script
 * Helps user identify all merged cells in a Google Sheet
 * 
 * To Do:
 * Identify merged cells across all tabs within Sheet
 * Output a summary of merged cell ranges and contents
 */

// highlight merged cells in specific sheet
function highlightMergedCells() {
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Sheet1');

  const range = sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns());
  const mergedRanges = range.getMergedRanges();
  
  mergedRanges.forEach(rng => {
    console.log(sheet.getName() + '!' + rng.getA1Notation()); // log merged cells range
    console.log(rng.getDisplayValue()); // log merged cells contents
    rng.setBackground('yellow');
  });

}
