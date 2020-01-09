/**
 * Create custom menu
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Color Tool')
        .addItem('Sort by color...', 'sortByColorSetupUi')
        .addItem('Clear Ranges','clearProperties')
        .addToUi();
}

/**
 * Sort By Color Setup Program Flow
 * Check whether color cell and sort columnn have been selected
 * If both selected, move to sort the data by color
 */
function sortByColorSetupUi() {
  
  var colorProperties = PropertiesService.getDocumentProperties();
  var colorCellRange = colorProperties.getProperty('colorCellRange');
  var sortColumnLetter = colorProperties.getProperty('sortColumnLetter');
  
  //if !colorCellRange
  if(!colorCellRange)  {
    title = 'Select Color Cell';
    msg = '<p>Please click on cell with the background color you want to sort on and then click OK</p>';
    msg += '<input type="button" value="OK" onclick="google.script.run.sortByColorHelper(1); google.script.host.close();" />';
    dispStatus(title, msg);
  }
  
  //if colorCellRange and !sortColumnLetter
  if (colorCellRange && !sortColumnLetter) {
      
      title = 'Select Sort Column';
      msg = '<p>Please highlight the column you want to sort on, or click on a cell in that column. Click OK when you are ready.</p>';
      msg += '<input type="button" value="OK" onclick="google.script.run.sortByColorHelper(2); google.script.host.close();" />';
      dispStatus(title, msg);
  }
  
  // both color cell and sort column selected
  if(colorCellRange && sortColumnLetter) {
    
    title= 'Displaying Color Cell and Sort Column Ranges';
    msg = '<p>Confirm ranges before sorting:</p>';
    msg += 'Color Cell Range: ' + colorCellRange + '<br />Sort Column: ' + sortColumnLetter + '<br />';
    msg += '<br /><input type="button" value="Sort By Color" onclick="google.script.run.sortData(); google.script.host.close();" />';
    msg += '<br /><br /><input type="button" value="Clear Choices and Exit" onclick="google.script.run.clearProperties(); google.script.host.close();" />';
    dispStatus(title,msg);
    
  }
}


/**
 * display the modeless dialog box
 */
function dispStatus(title,html) {
  
  var title = typeof(title) !== 'undefined' ? title : 'No Title Provided';
  var html = typeof(html) !== 'undefined' ? html : '<p>No html provided.</p>';
  var htmlOutput = HtmlService
     .createHtmlOutput(html)
     .setWidth(350)
     .setHeight(200);
 
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, title);

}


/**
 * helper function to switch between dialog box 1 (to select color cell) and 2 (to select sort column)
 */
function sortByColorHelper(mode) {
  
  var mode = (typeof(mode) !== 'undefined')? mode : 0;
  switch(mode)
  {
    case 1:
      setColorCell();
      sortByColorSetupUi();
      break;
    case 2:
      setSortColumn();
      sortByColorSetupUi();
      break;
    default:
      clearProperties();
  }
}

/** 
 * saves the color cell range to properties
 */
function setColorCell() {
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var colorCell = SpreadsheetApp.getActiveRange().getA1Notation();
  var colorProperties = PropertiesService.getDocumentProperties();
  colorProperties.setProperty('colorCellRange', colorCell);

}

/**
 * saves the sort column range in properties
 */
function setSortColumn() {
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var sortColumn = SpreadsheetApp.getActiveRange().getA1Notation();
  var sortColumnLetter = sortColumn.split(':')[0].replace(/\d/g,'').toUpperCase(); // find the column letter
  var colorProperties = PropertiesService.getDocumentProperties();
  colorProperties.setProperty('sortColumnLetter', sortColumnLetter);
  
}

/** 
 * sort the data based on color cell and chosen column
 */
function sortData() {
  
  // get the properties
  var colorProperties = PropertiesService.getDocumentProperties();
  var colorCell = colorProperties.getProperty('colorCellRange');
  var sortColumnLetter = colorProperties.getProperty('sortColumnLetter');
  
  // get the sheet
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  // get an array of background colors from the sort column
  var sortColBackgrounds = sheet.getRange(sortColumnLetter + 2 + ":" + sortColumnLetter + lastRow).getBackgrounds(); // assumes header in row 1
  
  // get the background color of the sort cell
  var sortColor = sheet.getRange(colorCell).getBackground();
  
  // map background colors to 1 if they match the sort cell color, 2 otherwise
  var sortCodes = sortColBackgrounds.map(function(val) {
    return (val[0] === sortColor) ? [1] : [2];
  });
  
  // add a column heading to the array of background colors
  sortCodes.unshift(['Sort Column']);
  
  // paste the background colors array as a helper column on right side of data
  sheet.getRange(1,lastCol+1,lastRow,1).setValues(sortCodes);
  sheet.getRange(1,lastCol+1,1,1).setHorizontalAlignment('center').setFontWeight('bold').setWrap(true);
  
  // sort the data
  var dataRange = sheet.getRange(2,1,lastRow,lastCol+1);  
  dataRange.sort(lastCol+1);
  
  // add new filter across whole data table
  sheet.getDataRange().createFilter();

  // clear out the properties so it's ready to run again
  clearProperties();
}

/**
 * clear the properties
 */
function clearProperties() {
  PropertiesService.getDocumentProperties().deleteAllProperties();
}