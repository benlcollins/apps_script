// on edit trigger
function onEdit(e) {
  
  // spreadsheet values
  const sheet = SpreadsheetApp.getActiveSheet();
  const val = sheet.getRange('A1').getValue();
  const thresholdVal = 100;

  // compare value to threshold
  if (val > thresholdVal) {
    // set background to yellow when value is greater than threshold
    sheet.getRange('A1').setBackground('yellow');
  }
  else {
    // set background to white when value is less than or equal to threshold
    sheet.getRange('A1').setBackground('white');
  }
}