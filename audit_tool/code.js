/** 
* Google Sheets Performance Auditor tool
* Built by Ben Collins, 2018
* https://www.benlcollins.com
*/


/**
* Custom menu
*/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sheet Size Auditor')
      .addItem('Audit Sheet', 'sheetAuditor')
      .addItem('Clear Data', 'clearData')
      .addToUi();
}


/**
* Clears out data and resets sheet
*/
function clearData() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  sheet.getRange(6,2).clearContent();
  sheet.getRange(11,1,1,5).clear(); 
  sheet.getRange(16,1,sheet.getLastRow(),13).clear();
  sheet.getRange(10,1,1,5).setBorder(true, true, true, true, true, true, "#D3D3D3", null); 
  sheet.getRange(15,1,1,13).setBorder(true, true, true, true, true, true, "#D3D3D3", null).setFontSize(14);  
}
  

/**
* Audits the given url and prints results into the Google Sheet table
*/
function sheetAuditor() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var url = sheet.getRange(6,2).getValue();
  
  try {
  
    var audit_data = auditUrl(url);
    
    // output the overall file data into my Google Sheet
    var fileRange = sheet.getRange(11,1,1,5);
    
    var fileFormats = [
      [ "#,###,##0", "#,###,##0", "0.0%", "#,###,##0", "0.0%" ]
    ];
    
    var fileHorizontalAlignments = [
      [ "center", "center", "center", "center", "center" ]
    ];
    
    fileRange.setValues([audit_data[0]]).setNumberFormats(fileFormats)
      .setHorizontalAlignments(fileHorizontalAlignments)
      .setBorder(true, true, true, true, true, true, "#D3D3D3", null)
      .setFontSize(14);
    
    // output the individual sheet data into my Google Sheet
    var sheetRange = sheet.getRange(16,1,audit_data[1].length,13);
    
    var sheetFormats = [
      [ "#,###,##0", "#,###,##0", "0.0%", "#,###,##0", "0.0%" ]
    ];
    
    sheetRange.setValues(audit_data[1]);
    
    var formatRange = sheet.getRange(16,1,audit_data[1].length,13);
    
    formatRange.setNumberFormat("#,###,##0")
      .setHorizontalAlignment("center")
      .setBorder(true, true, true, true, true, true, "#D3D3D3", null)
      .setFontSize(14);
  }
  catch(e) {
    Browser.msgBox("Please enter a valid Google Sheets URL");
  }
}


/**
* Returns performance and size data for a given sheet url
* @param {string} url - url of the Google Sheet
* @returns {array} of file data and individual sheets data
*/
function auditUrl(url) {
  
  var ss = SpreadsheetApp.openByUrl(url);
  var sheets = ss.getSheets();
  
  // counters
  var numSheets = 0;
  var totalCellCounter = 0;
  var totalDataCellCounter = 0;
  var fileArray = [];
  var sheetsArray = [];
  
  sheets.forEach(function(sheet) {

    // get single sheet data
    var thisSheetInfo = getSingleSheetInfo(sheet);
    
    // update the counters
    numSheets++;
    totalCellCounter = totalCellCounter + thisSheetInfo[3];
    totalDataCellCounter = totalDataCellCounter + thisSheetInfo[4];
    
    sheetsArray.push(thisSheetInfo);
  });
  
  var totalCellPercent = totalCellCounter / 5000000;
  var totalDataCellPercent = totalDataCellCounter / 5000000;
  
  fileArray.push(numSheets, totalCellCounter, totalCellPercent, totalDataCellCounter, totalDataCellPercent);
  
  return [fileArray,sheetsArray];
 
}


/**
* Returns individual sheet performance and size data
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
  
  // how many cells have data in them
  var r = sheet.getLastRow();
  var c = sheet.getLastColumn();
  var data_counter = r * c;
  
  if (data_counter !== 0) {
  
    var dataRange = sheet.getRange(1,1,r,c);
    var dataValues = dataRange.getValues();

    dataValues.forEach(function(row) {
      row.forEach(function(cell) {
        if (cell === "") {
          data_counter --;
        }
      });
    });  
  }
  
  // count how many volatile formulas
  var slowFuncs = slowFunctions(sheet);
    
  singleSheetArray.push(
    name,
    maxRows,
    maxCols,
    totalCells,
    data_counter, 
    slowFuncs[0], 
    slowFuncs[1], 
    slowFuncs[2], 
    slowFuncs[3], 
    slowFuncs[4],
    slowFuncs[5],
    slowFuncs[6],
    slowFuncs[7]
  );
  
  return singleSheetArray;
}


function slowFunctions(sheet) {
  
  var vols = identifyVolatiles(sheet);      //returns 7
  var arrs = identifyArrayFormulas(sheet);  //returns 1
  vols.push(arrs); 
  
  return vols;                              //returns 8
}


/**
* Returns counts for volatile functions in individual Google Sheet
* @param {object} sheet - the individual sheet within a Google Sheet
* @returns {array} of volatile function counts
*/
function identifyVolatiles(sheet) {
  
  // how many cells have data in them
  var r = sheet.getLastRow();
  var c = sheet.getLastColumn();
  var data_counter = r * c;
  
  var nowCounter = 0;
  var todayCounter = 0;
  var randCounter = 0;
  var randbetweenCounter = 0;
  var queryCounter = 0;  
  var iferrorCounter = 0;
  var indirectCounter = 0;
  var reNow = /.*NOW.*/;
  var reToday = /.*TODAY.*/;
  var reRand = /.*RAND.*/;
  var reRandbetween = /.*RANDBETWEEN.*/;
  var reQuery = /.*QUERY.*/;
  var reIfError = /.*IFERROR.*/;
  var reInDirect = /.*INDIRECT.*/;
  
  if (data_counter !== 0) {
  
    var dataRange = sheet.getRange(1,1,r,c);
    var formulaCells = dataRange.getFormulas();
    
    formulaCells.forEach(function(row) {
      row.forEach(function(cell) {
        if (cell.toUpperCase().match(reNow)) { nowCounter ++; };
        if (cell.toUpperCase().match(reToday)) { todayCounter++; };
        if (cell.toUpperCase().match(reRand) && !cell.toUpperCase().match(reRandbetween)) { randCounter++; };
        if (cell.toUpperCase().match(reRandbetween)) { randbetweenCounter++; }; 
        if (cell.toUpperCase().match(reQuery)) { 
          queryCounter++;  //setting to its own line for breakpoint debugging
        };   
        if (cell.toUpperCase().match(reIfError)) { iferrorCounter++; };   
        if (cell.toUpperCase().match(reInDirect)) { indirectCounter++; };   
      });
    });
  }
  
  return [nowCounter, todayCounter, randCounter, randbetweenCounter, queryCounter, iferrorCounter, indirectCounter];
  
}


/**
* Returns counts for ArrayFormula functions in individual Google Sheet
* @param {object} sheet - the individual sheet within a Google Sheet
* @returns {array} of ArrayFormula function counts
*/
function identifyArrayFormulas(sheet) {
  
  // how many cells have data in them
  var r = sheet.getLastRow();
  var c = sheet.getLastColumn();
  var data_counter = r * c;
  
  var arrayCounter = 0;
  var reArray = /.*ARRAYFORMULA.*/;
  
  if (data_counter !== 0) {
  
    var dataRange = sheet.getRange(1,1,r,c);    
    var formulaCells = dataRange.getFormulas();

    formulaCells.forEach(function(row) {
      row.forEach(function(cell) {
        if (cell.toUpperCase().match(reArray)) { 
          arrayCounter ++; 
        };
      });
    });    
  }
  
  return arrayCounter;

}
