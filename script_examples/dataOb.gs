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
  //Logger.log(headers);
  
  var headerIndexes = indexifyHeaders(headers);
  //Logger.log(JSON.stringify(headerIndexes));
  
  // check index still works
  allData.forEach(function(row) {
    //Logger.log(row[headerIndexes.carrier]);
    //Logger.log(row[headerIndexes.name]);
    //Logger.log(row[headerIndexes["annual flights"]]);
  });
  
  // insert a new column 
  newSheet.insertColumnBefore(headerIndexes.name + 1);
  
  // reget the data
  var allData = newSheet.getDataRange().getValues();
  
  // extract the new header row
  var headers = allData.shift();
  
  // make a new index header map
  var headerIndexes = indexifyHeaders(headers);
  //Logger.log(headerIndexes);
  
  // check index still works
  allData.forEach(function(row) {
    //Logger.log(row[headerIndexes.carrier]);
    //Logger.log(row[headerIndexes.name]);
    //Logger.log(row[headerIndexes["annual flights"]]);
  });
  
  // convert all the data to an object
  var dataAsObjects = objectifyData (headerIndexes , allData);
  //Logger.log(JSON.stringify(dataAsObjects));
  
  
  // combine both
  var shortcutData = objectifyRange(newSheet.getDataRange());
  Logger.log(JSON.stringify(shortcutData));
  
  // clear the sheet
  newSheet.clear();
  
  // write back modified data (minus the blank column)
  var newData = datifyObjects(shortcutData.filter(function(row) {
    return row['annual flights'] > 4500;
  }));
  
  // write out the data
  newSheet.getRange(
    1,1,newData.length,newData[0].length
  ).setValues(newData);
  
}

/**
 * turn objectified data into sheet writable data
 * @param {[object]} dataObjects an array of objectified sheet data
 * @return {[[]]} a data array in sheet format
 */
function datifyObjects (dataObjects) {
  
  // get the headers from a row of objects
  var headers = Object.keys(dataObjects[0]);
  
  // turn the data back to an array and concat to header
  return [headers].concat(dataObjects.map(function(row) {
    return headers.map(function(cell) {
      return row[cell];
    });
  }));
}





/**
 * create an object from a range input
 * @param {[*]} range a range of data
 * @return {object} an object where the props are names & values are indexes
 */
function objectifyRange(range) {
  
  var allData = range.getValues();
  //Logger.log(allData);
  
  // [[carrier, , name, annual flights], [8P, , Pacific Air Coast, 4603.0], [9R, , Satena, 9849.0], [AC, , Air Canada, 663.0], [AY, , Finnair, 4083.0]]
  
  // extract the new header row
  var headers = allData.shift();
  
  // make a new index header map
  var headerIndexes = indexifyHeaders(headers);
  
  // convert all the data to an object and return
  return objectifyData (headerIndexes , allData);
  
}


//  [{"carrier":"8P","name":"Pacific Air Coast","annual flights":4603},{"carrier":"9R","name":"Satena","annual flights":9849},
// {"carrier":"AC","name":"Air Canada","annual flights":663},{"carrier":"AY","name":"Finnair","annual flights":4083}]




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
  
  
  
  
  
