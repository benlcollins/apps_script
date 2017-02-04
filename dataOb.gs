// From screencast 24 of Bruce McPherson's Apps Script Developer course
// http://shop.oreilly.com/product/0636920048503.do

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
  Logger.log(JSON.stringify(headerIndexes));
  
  // check index still works
  allData.forEach(function(row) {
    Logger.log(row[headerIndexes.carrier]);
    Logger.log(row[headerIndexes.name]);
    Logger.log(row[headerIndexes["annual flights"]]);
  });
  
  // insert a new column 
  newSheet.insertColumnBefore(headerIndexes.name + 1);
  
  // reget the data
  var allData = newSheet.getDataRange().getValues();
  
  // extract the new header row
  var headers = allData.shift();
  
  // make a new index header map
  var headerIndexes = indexifyHeaders(headers);
  Logger.log(headerIndexes);
  
  // check index still works
  allData.forEach(function(row) {
    Logger.log(row[headerIndexes.carrier]);
    Logger.log(row[headerIndexes.name]);
    Logger.log(row[headerIndexes["annual flights"]]);
  });
  
  // convert all the data to an object
  var dataAsObjects = objectifyData (headerIndexes , allData);
  Logger.log(JSON.stringify(dataAsObjects));
  
  
  
  
}


/**
 * create a map of indexes to properties
 * @param {[*]} headers an array of header names 
 * @return {object} an object where the props are names & values are indexes
 */
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
  

/**
 * create an array of objects from data
 * @param {object} headerIndexes the map of header names to column indexes
 * @param {[[]]} data the data from the sheet
 * @return {[object]} the objectified data
 */
function objectifyData (headerIndexes, data) {
  return data.map(function(row) {
    return Object.keys(headerIndexes).reduce (function(p,c) {
      p[c] = row[headerIndexes[c]];
      return p;
    },{});
  });
}
  
  
  
  
  
