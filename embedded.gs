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
  
  var data = range.getValues();
  
  var newData = [['Label','Base','Endpoints','Postive Cols above','Positive Cols below',
                  'Negative Cols above','Negative Cols below','Cross axis positive above',
                  'Cross axis positive below','Cross axis negative above','Cross axis negative below']];
  
  var tempTotal = 0;
  var tempTotalPrior = 0;
  var tempObj = {};
  
  for (var i = 1; i < data.length; i++) {
    // add up total of values so far
    tempTotalPrior = tempTotal;
    tempTotal += data[i][1];
    
    // create temp object with running totals and value
    tempObj["position"] = i;
    tempObj["label"] = data[i][0];
    tempObj["value"] = data[i][1];
    tempObj["tempTotal"] = tempTotal;
    tempObj["tempTotalPrior"] = tempTotalPrior;
    //tempObj["valueAbs"] = Math.abs(data[i][1]);
    //tempObj["tempTotalAbs"] = Math.abs(tempTotal);
    //tempObj["tempTotalPriorAbs"] = Math.abs(tempTotalPrior);
    
    // Endpoints
    if (tempObj.position == 1 || tempObj.position == data.length - 1) {
      newData.push([tempObj.label,0,tempObj.value,'','','','','','','','']);
    }
    
    // Cross axis positive
    else if (tempObj.value > 0 && tempObj.tempTotal > 0 && tempObj.tempTotalPrior < 0) {
      newData.push([tempObj.label,0,'','','','','',tempObj.tempTotal,tempObj.tempTotalPrior,'','']);
    }
    
    // Cross axis negative
    else if (tempObj.value < 0 && tempObj.tempTotal < 0 && tempObj.tempTotalPrior > 0) {
      newData.push([tempObj.label,0,'','','','','','','',tempObj.tempTotalPrior,tempObj.tempTotal]);
    }    

    // Postive Cols above
    else if (tempObj.value > 0 && tempObj.tempTotalPrior > 0) {
      newData.push([tempObj.label,tempTotalPrior,'',tempObj.value,'','','','','','','']);
    }
    
    // Postive Cols below
    else if (tempObj.value > 0 && tempObj.tempTotalPrior < 0) {
      newData.push([tempObj.label,tempTotal,'','',-tempObj.value,'','','','','','']);
    }
    
    // Negative Cols above
    else if (tempObj.value < 0 && tempObj.tempTotalPrior > 0) {
      newData.push([tempObj.label,tempTotal,'','','',-tempObj.value,'','','','','']);
    }
    
    // Negative Cols below
    else if (tempObj.value < 0 && tempObj.tempTotalPrior < 0) {
      newData.push([tempObj.label,tempTotalPrior,'','','','',tempObj.value,'','','','']);
    }
    
    // in case of no data
    else {
      newData.push(["Error",i,tempObj.label,tempObj.value,'','','','','','','']);
    }
    
  }
  
  // paste the new data into sheet
  sheet.getRange(1, 4, data.length, newData[0].length).setValues(newData);
  
  // get the new data for the chart
  var chartData = sheet.getRange(1, 4, data.length, newData[0].length);
  
  // make the new funnel chart
  sheet.insertChart(
    sheet.newChart()
    .addRange(chartData)
    .setChartType(Charts.ChartType.COLUMN)
    .asColumnChart()
    .setStacked()
    .setOption('title','Waterfall chart')
    .setLegendPosition(Charts.Position.NONE)
    .setPosition(11,4,0,0)
    .build()
  );
}