//add menu to google sheet
function onOpen() {
  //set up custom menu
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Waterfall Chart')
    .addItem('Insert chart...','waterfallChart')
    .addToUi();
};

// function to create waterfall chart
function waterfallChart() {
  
  // get the sheet
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // get the range highlighted by user
  var range = sheet.getActiveRange();
  var data = range.getValues();
  
  var newData = [['Label','Endpoints','Base','Postive Cols above','Positive Cols below',
                  'Negative Cols above','Negative Cols below']];
  
  var tempTotal = 0;
  var tempTotalPrior = 0;
  
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  for (i = 1; i < data.length; i++) {
    
    // running totals
    tempTotalPrior = tempTotal;  // assign previous total to new variable to keep track of it
    tempTotal += data[i][1];  // add up total of values so far
    
    // Endpoints
    if (i == 1 || i == data.length - 1) {
      newData.push([data[i][0],data[i][1],0,'','','','']);
    }
    
    // Non-endpoints
    else {
      
      // Base values
      var baseVal = Math.max(0,Math.min(tempTotal,tempTotalPrior)) + Math.min(0,Math.max(tempTotal,tempTotalPrior));
      
      // calculate minimum of running total and current value
      var val1 = Math.min(tempTotal,data[i][1]);
      
      // calculate maximum of running total and current value
      var val2 = Math.max(tempTotal,data[i][1]);
      
      // Postive Cols above
      // if val1 is negative, set to 0, otherwise take val1, which is min of running total and current value
      var posValAbove = Math.max(0,val1);
      
      // Postive Cols below
      // subtract current value from Positive Col Value to catch any part of column below 0. If a positive value, set to 0 by using minimum
      var posValBelow = Math.min(posValAbove - data[i][1],0);
      
      // Negative Cols below
      // if val2 is positive, set to 0, otherwise take val2, which is the max of running total and current value
      var negValBelow = Math.min(0,val2);
      
      // Negative Cols above
      // subtract current value from Negative Col Value to catch any part of column above 0. If a negative value, set to 0 by using maximum
      var negValAbove = Math.max(negValBelow - data[i][1],0);
       
      // push all new datapoints into newData array
      newData.push([data[i][0],0,baseVal,posValAbove,posValBelow,negValAbove,negValBelow]);
    }
    
  }
  
  // paste the new data into sheet
  sheet.getRange(lastRow - data.length + 1, lastCol + 2, data.length, newData[0].length).setValues(newData);

  // get the new data for the chart
  var chartData = sheet.getRange(lastRow - data.length + 1, lastCol + 2, data.length, newData[0].length);
  
  // make the new waterfall chart
  sheet.insertChart(
    sheet.newChart()
    .addRange(chartData)
    .setChartType(Charts.ChartType.COLUMN)
    .asColumnChart()
    .setStacked()
    .setOption('title','Waterfall Chart')
    .setLegendPosition(Charts.Position.NONE)
    .setPosition(lastRow - data.length + 4,lastCol + 4,0,0)
    .build()
  );
}