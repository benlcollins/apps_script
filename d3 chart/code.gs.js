function doGet(e) {
  return HtmlService.createTemplateFromFile("index.html")
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// get the chart data to pass through to front-end
function getChartData() {
  
  var ss = SpreadsheetApp.openById("XXXXXXXXXXXXXX");
  var sheet = ss.getActiveSheet();
  
  var headings = sheet.getRange(3,1,1,sheet.getLastColumn()).getValues()[0].map(function(heading) {
    return heading.toLowerCase();
  });
  
  Logger.log(headings);
  
  var values = sheet.getRange(4, 1, sheet.getLastRow()-3, sheet.getLastColumn()).getValues();
  
  var data = [];
  for (var i=0; i < values.length; i++) {
    var obj = {};
    for (var j = 0; j < values[i].length; j++) {
      obj[headings[j]] = values[i][j];
    }
    data.push(obj);
  }
  
  return data;
}


