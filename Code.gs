//add menu to google sheet
function onOpen() {
  //set up custom menu
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Waterfall Chart')
    .addItem('Waterfall chart','waterfallChart')
    .addToUi();
};


// function to create waterfall chart
function waterfallChart() {
  
  // get the sheet
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // get the range highlighted by user
  var range = sheet.getDataRange();
  
  var data = range.getValues();
  
  
  
}
