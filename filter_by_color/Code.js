/**
 * Create custom menu
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Color Tool')
        .addItem('Filter by color...', 'filterByColorSetupUi')
        .addItem('Clear Ranges','clearProperties')
        .addToUi();
}

/**
 * Filter By Color Setup Program Flow
 * Check whether color cell and filter columnn have been selected
 * If both selected, move to filter the data by color
 */
function filterByColorSetupUi() {
  
  var colorProperties = PropertiesService.getDocumentProperties();
  var colorCellRange = colorProperties.getProperty('colorCellRange');
  var filterColumnLetter = colorProperties.getProperty('filterColumnLetter');
  
  //if !colorCellRange
  if(!colorCellRange)  {
    title = 'Select Color Cell';
    msg = '<p>Please click on cell with the background color you want to filter on and then click OK</p>';
    msg += '<input type="button" value="OK" onclick="google.script.run.filterByColorHelper(1); google.script.host.close();" />';
    dispStatus(title, msg);
  }
  
  //if colorCellRange and !filterColumnLetter
  if (colorCellRange && !filterColumnLetter) {
      
      title = 'Select Filter Column';
      msg = '<p>Please highlight the column you want to filter, or click on a cell in that column. Click OK when you are ready.</p>';
      msg += '<input type="button" value="OK" onclick="google.script.run.filterByColorHelper(2); google.script.host.close();" />';
      dispStatus(title, msg);
  }
  
  // both color cell and filter column selected
  if(colorCellRange && filterColumnLetter) {
    
    title= 'Displaying Color Cell and Filter Column Ranges';
    msg = '<p>Confirm ranges before filtering:</p>';
    msg += 'Color Cell Range: ' + colorCellRange + '<br />Filter Column: ' + filterColumnLetter + '<br />';
    msg += '<br /><input type="button" value="Filter By Color" onclick="google.script.run.filterData(); google.script.host.close();" />';
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
 * helper function to switch between dialog box 1 (to select color cell) and 2 (to select filter column)
 */
function filterByColorHelper(mode) {
  
  var mode = (typeof(mode) !== 'undefined')? mode : 0;
  switch(mode)
  {
    case 1:
      setColorCell();
      filterByColorSetupUi();
      break;
    case 2:
      setFilterColumn();
      filterByColorSetupUi();
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
 * saves the filter column range in properties
 */
function setFilterColumn() {
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var filterColumn = SpreadsheetApp.getActiveRange().getA1Notation(); 
  var filterColumnLetter = filterColumn.split(':')[0].replace(/\d/g,'').toUpperCase(); // extracts column letter from whatever range has been highlighted for the filter column
  var colorProperties = PropertiesService.getDocumentProperties();
  colorProperties.setProperty('filterColumnLetter', filterColumnLetter);
  
}

/** 
 * filter the data based on color cell and chosen column
 */
function filterData() {
  
  // get the properties
  var colorProperties = PropertiesService.getDocumentProperties();
  var colorCell = colorProperties.getProperty('colorCellRange');
  var filterColumnLetter = colorProperties.getProperty('filterColumnLetter');
  
  // get the sheet
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  // get an array of background colors from the filter column
  var filterColBackgrounds = sheet.getRange(filterColumnLetter + 2 + ":" + filterColumnLetter + lastRow).getBackgrounds(); // assumes header in row 1
  
  // add a column heading to the array of background colors
  filterColBackgrounds.unshift(['Column ' + filterColumnLetter + ' background colors']);
  
  // paste the background colors array as a helper column on right side of data
  sheet.getRange(1,lastCol+1,lastRow,1).setValues(filterColBackgrounds);
  sheet.getRange(1,lastCol+1,1,1).setHorizontalAlignment('center').setFontWeight('bold').setWrap(true);
  
  // get the background color of the filter cell
  var filterColor = sheet.getRange(colorCell).getBackground();
  
  // remove existing filter to the data range
  if (sheet.getFilter() !== null) {
    sheet.getFilter().remove();
  }
  
  // add new filter across whole data table
  var newFilter = sheet.getDataRange().createFilter();
  
  // create new filter criteria
  var filterCriteria = SpreadsheetApp.newFilterCriteria();
  filterCriteria.whenTextEqualTo(filterColor);
  
  // apply the filter color as the filter value
  newFilter.setColumnFilterCriteria(lastCol + 1, filterCriteria);
  
  // clear out the properties so it's ready to run again
  clearProperties();
}

/**
 * clear the properties
 */
function clearProperties() {
  PropertiesService.getDocumentProperties().deleteAllProperties();
}