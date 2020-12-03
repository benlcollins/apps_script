// add shortcut menu to sheet
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('RGB Menu')
      .addItem('Show color', 'rgbColor')
      .addToUi();
}

// function to set RGB color on row 2
function rgbColor() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const rVal = sheet.getRange(2,1).getValue();
  const gVal = sheet.getRange(2,2).getValue();
  const bVal = sheet.getRange(2,3).getValue();
  sheet.getRange(2,4).setBackgroundRGB(rVal,gVal,bVal);
}