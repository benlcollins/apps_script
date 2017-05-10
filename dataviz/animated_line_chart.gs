function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu("Timed data")
    .addItem("Start","startTimedData")
    .addItem("Clear","clearData")
    .addToUi();
}
  

function startTimedData() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Animated Chart');
  var lastRow = sheet.getLastRow()-12;
  
  var data2015 = sheet.getRange(13,2,lastRow,1).getValues(); // historic data
  var data2016 = sheet.getRange(13,5,lastRow,1).getValues(); // historic data
  
  // new data that would be inputted into the sheet manually or from API
  var data2017 = [[1],[7],[14],[19],[27],[32],[34],[36],[44],[49],[57],[65],[72],[76],[79],[86],[92],[99],[104],[109],[111],[112],[120],[128],[130],
                  [132],[133],[140],[144],[149],[151],[152],[158],[162],[170],[177],[179],[184],[188],[194],[200],[205],[211],[216],[224],[232],[238],
                  [241],[246],[248],[252],[259],[266],[268],[276],[284],[291],[299],[300],[301],[306],[311],[315],[316],[323],[324]];
  
  for (var i = 0; i < data2015.length;i++) {
    outputData(data2015[i],data2016[i],data2017[i],i);
  }
  
}

function outputData(d1,d2,d3,i) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Animated Chart');
  
  sheet.getRange(13+i,3).setValue(d1);
  sheet.getRange(13+i,6).setValue(d2);
  sheet.getRange(13+i,8).setValue(d3);
  Utilities.sleep(10);
  SpreadsheetApp.flush();
}

function clearData() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Animated Chart');
  var lastRow = sheet.getLastRow()-12;
  
  sheet.getRange(13,3,lastRow,1).clear();
  sheet.getRange(13,6,lastRow,1).clear();
  sheet.getRange(13,8,lastRow,1).clear();
  
}
