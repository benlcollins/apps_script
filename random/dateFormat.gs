function dateFormat() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Sheet1');
  
  var data = sheet.getRange(1,14,sheet.getLastRow(),1).getValues();
  
  var newDataArray = [];
  
  data.forEach(function(el) {
    var cleanDate = el[0].replace("Date(","").replace(")","").split(",");
    if (cleanDate[0]) {
      newDataArray.push([cleanDate[1] + "/" + cleanDate[2] + "/" + cleanDate[0]]);
    }
    else {
      newDataArray.push([""]);
    }
  });
  
  sheet.getRange(1,14,sheet.getLastRow(),1).setValues(newDataArray);
}