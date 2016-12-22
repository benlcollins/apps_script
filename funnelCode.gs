//add menu to google sheet
function onOpen() {
  
  //set up custom menu
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Chart')
    .addItem('Funnel chart','funnelChart')
    .addItem('Sparkline Funnel chart','sparklineFunnelChart')
    .addItem('REPT Funnel chart','reptFunnelChart')
    .addToUi();
};

// sort the data first
function dataSort() {
  // TBC
};

// embedded chart builder
function funnelChart() {
  
  // get the sheet
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // get the range highlighted by user
  var range = sheet.getActiveRange();
  var data = range.getValues();
  
  // move the values one column to the right, to make space for helper column
  range.offset(0,range.getNumColumns(),range.getNumRows(),1)
    .setValues(data.map(function(d){
      return [d[1]];
  }));
  
  // get the max of the data
  var maxVal = Math.max.apply(Math, data.map(function(d){
      return [d[1]];
  }));
  
  // create the helper column
  range.offset(0,range.getNumColumns()-1,range.getNumRows(),1)
    .setValues(data.map(function(d){
      return [(maxVal - d[1]) / 2];
  }));
  
  // make the new funnel chart
  // To do: add the data labels to the bars with the annotations option
  // see: https://developers.google.com/chart/interactive/docs/gallery/barchart#labeling-bars
  sheet.insertChart(
    sheet.newChart()
    .addRange(sheet.getDataRange())
    .setChartType(Charts.ChartType.BAR)
    .asBarChart()
    .setColors(["none", "#FFA500"])
    .setStacked()
    .setOption("title","GAS Funnel chart")
    .setOption("hAxis.gridlines.color","none")
    .setOption("hAxis.textStyle.color","none")
    .setLegendPosition(Charts.Position.NONE)
    .setPosition(2, sheet.getLastColumn()+2,0,0)
    .build()
  );
};


// sparkline chart builder
function sparklineFunnelChart() {
  
  // get the sheet
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // get the range and data highlighted by user
  var range = sheet.getActiveRange();  
  var data = range.getValues();
  var len = data.length;
  
  // sort the range by the second column of values, highest to lowest and find max value
  var sortedRange = range.sort({column: range.getColumn()+1, ascending: false});
  var maxVal = sortedRange.getValues()[0][1];
  
  // create array to hold sparklines left
  var formulasLeft = [];
  for (j = 0; j < len; j++) {
    var optionsLeft = '{"charttype","bar";"max",R[-' + j + ']C[-2]; "rtl",true; "color1","#FFA500"}';
    var sparklineFunnelLeft = "=SPARKLINE(R[0]C[-2]," + optionsLeft + ")";
    formulasLeft.push([sparklineFunnelLeft]);
  }

  // create array to hold sparklines right
  var formulasRight = [];
  for (i = 0; i < len; i++) {
    var optionsRight = '{"charttype","bar";"max",R[-' + i + ']C[-3]; "color1","#FFA500"}';
    var sparklineFunnelRight = "=SPARKLINE(R[0]C[-3]," + optionsRight + ")";
    formulasRight.push([sparklineFunnelRight]);
  }
  
  // identify range for sparkline output
  var funnelOutputLeft = sheet.getRange(range.getRowIndex(),range.getColumn()+3,len,1);
  var funnelOutputRight = sheet.getRange(range.getRowIndex(),range.getColumn()+4,len,1);
  
  // put the sparkline formulas into the output ranges
  funnelOutputLeft.setFormulasR1C1(formulasLeft);
  funnelOutputRight.setFormulasR1C1(formulasRight);
};


// REPT chart builder
function reptFunnelChart() {
  
  // get the sheet
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // get the range and data highlighted by user
  var range = sheet.getActiveRange();  
  var data = range.getValues();
  var len = data.length;
  
  // sort the range by the second column of values, highest to lowest and find max value
  var sortedRange = range.sort({column: range.getColumn()+1, ascending: false});
  var maxVal = sortedRange.getValues()[0][1];
  var scaleFactor = maxVal / 65; // trial and error to calculate that this gives a good scaling for a column width of 250
  
  // create array to hold REPT formulas
  var formulas = [];
  for (j = 0; j < len; j++) {
    var rept = '=REPT("|",R[0]C[-2] /' + scaleFactor +')';
    formulas.push([rept]);
  }
  
  // identify range for REPT output
  var reptOutput = sheet.getRange(range.getRowIndex(),range.getColumn()+3,len,1);
  
  // format output range for REPT funnel chart
  reptOutput.setFontFamily("Modak");
  reptOutput.setFontColor("#FFA500");
  reptOutput.setHorizontalAlignment("center");
  
  // set the column width to 250 for the REPT chart column
  sheet.setColumnWidth(range.getColumn()+3, 250)
  
  // put the sparkline formulas into the output ranges
  reptOutput.setFormulasR1C1(formulas);
};
