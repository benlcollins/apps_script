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
  var range = sheet.getActiveRange();
  //Logger.log(range);
  
  var data = range.getValues();
  Logger.log(data);
  // [[Label, Value $], [Start Loss, -20.0], [Revenue, 65.0], [Cost of Sales, -50.0], [Salary costs, -20.0], [Other costs, -10.0], [Other income, 5.0], [End Loss, -30.0]]
  
  Logger.log(data.length);
  
  // [[Label, Value $, Value $], [Start Loss, , -20.0], [Revenue, 65.0, ], [Cost of Sales, , -50.0], [Salary costs, , -20.0], [Other costs, , -10.0], [Other income, 5.0, ], [End Loss, , -30.0]]
  
  var newData = [['Label','Base','Endpoints','Postive Cols above','Positive Cols below',
                  'Negative Cols above','Negative Cols below','Cross axis positive above',
                  'Cross axis positive below','Cross axis negative above','Cross axis negative below']];
  
  for (var i = 0; i < data.length; i++) {
    newData.push([data[i][0],data[i][1],data[i][2],0,0,0,0,0,0,0,0]);
  }
  
  Logger.log(newData);
  
  //var chartData = google.visualization.arrayToDataTable([
  //[]]);
  

  /*
  //make the new funnel chart
  sheet.insertChart(
    sheet.newChart()
    .addRange(range)
    .setChartType(Charts.ChartType.COLUMN)
    .asColumnChart()
    .setStacked()
    .setOption('title','Waterfall chart')
    //.setColors('#ff0000','#ff00ff')
    .setLegendPosition(Charts.Position.NONE)
    .setPosition(4, 4,0,0)
    .build()
  );
  */
}
