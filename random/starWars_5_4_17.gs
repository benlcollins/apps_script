function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Star Wars')
    .addItem('Star Wars','starWarsStory')
    .addToUi(); 
}

function starWarsStory() {
 for (var i = 0; i < 10; i++) {
    copyData(2+i*14);
  } 
}

function copyData(r) {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('starWars');
  var storyCells = sheet.getRange(2, 1, 13, 9);
  
  var range = sheet.getRange(r, 11, 13, 9).copyTo(storyCells);
  Utilities.sleep(1000);
  SpreadsheetApp.flush();
}
