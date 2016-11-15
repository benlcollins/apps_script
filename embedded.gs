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
  //Logger.log(data);
  // [[Label, Value $], [Start Loss, -20.0], [Revenue, 65.0], [Cost of Sales, -50.0], [Salary costs, -20.0], [Other costs, -10.0], [Other income, 5.0], [End Loss, -30.0]]
  
  //Logger.log(data.length);
  
  // [[Label, Value $, Value $], [Start Loss, , -20.0], [Revenue, 65.0, ], [Cost of Sales, , -50.0], [Salary costs, , -20.0], [Other costs, , -10.0], [Other income, 5.0, ], [End Loss, , -30.0]]
  
  var newData = [['Label','Base','Endpoints','Postive Cols above','Positive Cols below',
                  'Negative Cols above','Negative Cols below','Cross axis positive above',
                  'Cross axis positive below','Cross axis negative above','Cross axis negative below']];
  
  var tempTotal = 0;
  
  for (var i = 1; i < data.length; i++) {
    // add up total of values so far
    tempTotal += data[i][1];
    Logger.log(tempTotal);
    
    if (i == 1 || i == data.length - 1) { // Endpoints
      newData.push([data[i][0],0,data[i][1],'','','','','','','','']);
    }
    
    // Postive Cols above
    
    else if (data[i][1] >= 0 && tempTotal < 0) { // Postive Cols below
      newData.push([data[i][0],tempTotal,'','',-data[i][1],'','','','','','']);
    }
    
    // Negative Cols above
    
    
    else if (data[i][1] < 0 && tempTotal < 0 && (tempTotal-data[i][1]) < 0) { // Negative Cols below
      newData.push([data[i][0],tempTotal-data[i][1],'','','','',data[i][1],'','','','']);
    }
    else if (tempTotal > 0 && data[i-1][1] < 0){  // Cross axis positive
      newData.push([data[i][0],0,'','','','','',tempTotal,data[i-1][1],'','']);
    }
    else if (tempTotal < 0 && data[i-1][1] > 0) {  // Cross axis negative
      newData.push([data[i][0],0,'','','','','','','',tempTotal-data[i][1],tempTotal]);
    }
    else {
      newData.push([data[i][0],'','',data[i][1],0,0,0,0,0,0,0]);
    }
  }
  
  sheet.getRange(1, 4, data.length, 11).setValues(newData);
  
  var chartData = sheet.getRange(1, 4, data.length, 11);
  
  // make the new funnel chart
  sheet.insertChart(
    sheet.newChart()
    .addRange(chartData)
    .setChartType(Charts.ChartType.COLUMN)
    .asColumnChart()
    .setStacked()
    .setOption('title','Waterfall chart')
    //.setColors('#ff0000','#ff00ff')
    .setLegendPosition(Charts.Position.NONE)
    .setPosition(11,4,0,0)
    .build()
  );
}


// function to create waterfall chart
function waterfallChart2() {
  
  // get the sheet
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // get the range highlighted by user
  var range = sheet.getActiveRange();
  
  var data = range.getValues();
  
  var newData = [['Label','Base','Endpoints','Postive Cols above','Positive Cols below',
                  'Negative Cols above','Negative Cols below','Cross axis positive above',
                  'Cross axis positive below','Cross axis negative above','Cross axis negative below']];
  
  var tempTotal = 0;
  var tempTotalPrior = 0;
  var tempObj = {};
  
  // obj["key3"] = "value3";
  
  for (var i = 1; i < data.length; i++) {
    // add up total of values so far
    tempTotalPrior = tempTotal;
    tempTotal += data[i][1];
    tempObj["Position"] = i;
    tempObj["Value"] = data[i][1];
    tempObj["tempTotal"] = tempTotal;
    tempObj["tempTotalPrior"] = tempTotalPrior;
    tempObj["Value_Abs"] = Math.abs(data[i][1]);
    tempObj["tempTotal_Abs"] = Math.abs(tempTotal);
    tempObj["tempTotalPrior_Abs"] = Math.abs(tempTotalPrior);
    
    //= [i,data[i][1],tempTotal,tempTotalPrior,Math.abs(data[i][1]),Math.abs(tempTotal),Math.abs(tempTotalPrior)];
    Logger.log(tempObj);
    
    // Endpoints
    if (tempObj.Position == 1 || tempObj.Position == data.length - 1) {
      newData.push([data[i][0],0,tempObj.Value,'','','','','','','','']);
    }
    
    // Cross axis positive
    
    // Cross axis negative
    

    // Postive Cols above
    else if (tempObj.Value > 0 && tempObj.tempTotalPrior > 0) {
      newData.push([data[i][0],tempTotalPrior,'',tempObj.Value,'','','','','','','']);
    }
    
    // Postive Cols below
    else if (tempObj.Value > 0 && tempObj.tempTotalPrior < 0) {
      newData.push([data[i][0],tempTotal,'','',tempObj.Value_Abs,'','','','','','']);
    }
    
    // Negative Cols above
    else if (tempObj.Value < 0 && tempObj.tempTotalPrior > 0) {
      newData.push([data[i][0],tempTotal,'','','',tempObj.Value_Abs,'','','','','']);
    }
    
    // Negative Cols below
    else if (tempObj.Value < 0 && tempObj.tempTotalPrior < 0) {
      newData.push([data[i][0],tempTotalPrior,'','','','',tempObj.Value,'','','','']);
    }
    
    
    else {
      newData.push([99,99,99,99,99,99,99,99,99,99,99]);
    }

    
  }
  
  Logger.log("New Data:");
  Logger.log(newData);
  // [[Label, Base, Endpoints, Postive Cols above, Positive Cols below, Negative Cols above, Negative Cols below, Cross axis positive above, Cross axis positive below, Cross axis negative above, Cross axis negative below], 
  // [Start Loss, 0.0, -20.0, , , , , , , , ], 
  // [Revenue, 45.0, , , -65.0, , , , , , ], 
  // [99.0, 99.0, 99.0, 99.0, 99.0, 99.0, 99.0, 99.0, 99.0, 99.0, 99.0], 
  // [Salary costs, -5.0, , , , , -20.0, , , ], 
  // [Other costs, -25.0, , , , , -10.0, , , ], 
  // [Other income, -30.0, , , -5.0, , , , , , ], 
  // [End Loss, 0.0, -30.0, , , , , , , , ]]
  
  sheet.getRange(1, 4, data.length, 11).setValues(newData);
  
}