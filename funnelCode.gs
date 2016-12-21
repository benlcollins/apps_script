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

// embedded chart builder
function funnelChart() {
  
  // get the sheet
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // get the range highlighted by user
  var range = sheet.getActiveRange();
  var data = range.getValues();
  
  range.offset(0,range.getNumColumns(),range.getNumRows(),1)
    .setValues(data.map(function(d){
      return [d[1]];
  }));
  
  // get the max of the data
  var maxVal = Math.max.apply(Math, data.map(function(d){
      return [d[1]];
  }));
  
  
  range.offset(0,range.getNumColumns()-1,range.getNumRows(),1)
    .setValues(data.map(function(d){
      return [(maxVal - d[1]) / 2];
  }));
  
  //make the new funnel chart
  sheet.insertChart(
    sheet.newChart()
    .addRange(sheet.getDataRange())
    .setChartType(Charts.ChartType.BAR)
    .asBarChart()
    .setColors(["none", "red"])
    .setStacked()
    .setOption('title','GAS Funnel chart')
    .setLegendPosition(Charts.Position.NONE)
    .setPosition(2, sheet.getLastColumn()+3,0,0)
    .build()
  );
};


// sparkline chart builder
function sparklineFunnelChart() {
  
  // get the sheet
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // get the range highlighted by user
  var range = sheet.getActiveRange();
  var data = range.getValues();
  var len = data.length;
  var maxVal = 636;
  
  var optionsLeft = '{"charttype","bar";"max",' + maxVal + '; "rtl",true; "color1","red"}';
  var optionsRight = '{"charttype","bar";"max",' + maxVal + '; "color1","red"}';
  
  //var sparklineFunnel = "=SPARKLINE(R[0]C[-2],{'charttype','bar';'max',max(R[0]C[-2]:R[maxVal]C[-2]); 'rtl',true;'color1','#FFA500'})";
  var sparklineFunnelLeft = "=SPARKLINE(R[0]C[-2]," + optionsLeft + ")";
  
  // array to hold sparklines right
  var formulasLeft = [];
  for (i = 0; i < len; i++) {
    formulasLeft.push([sparklineFunnelLeft]);
    
  }
  
  //var sparklineFunnel = "=SPARKLINE(R[0]C[-2],{'charttype','bar';'max',max(R[0]C[-2]:R[maxVal]C[-2]); 'rtl',true;'color1','#FFA500'})";
  var sparklineFunnelRight = "=SPARKLINE(R[0]C[-3]," + optionsRight + ")";
  
  // array to hold sparklines right
  var formulasRight = [];
  for (i = 0; i < len; i++) {
    formulasRight.push([sparklineFunnelRight]);
    
  }
  
  // Logger.log(formulas);
  
  var funnelOutputLeft = sheet.getRange(range.getRowIndex(),range.getColumn()+3,len,1);
  var funnelOutputRight = sheet.getRange(range.getRowIndex(),range.getColumn()+4,len,1);
  
  funnelOutputLeft.setFormulasR1C1(formulasLeft);
  funnelOutputRight.setFormulasR1C1(formulasRight);
};


// sparkline chart builder
function reptFunnelChart() {
  
  // get the sheet
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // get the range highlighted by user
  var range = sheet.getActiveRange();
  var data = range.getValues();
  Logger.log(data);
};

// test
function testR1C1() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getSheets()[0];

 // This creates formulas for a row of sums, followed by a row of averages.
 var sumOfRowsAbove = "=SUM(R[-3]C[0]:R[-1]C[0])";
 var averageOfRowsAbove = "=AVERAGE(R[-4]C[0]:R[-2]C[0])";

 // The size of the two-dimensional array must match the size of the range.
 var formulas = [
   [sumOfRowsAbove, sumOfRowsAbove, sumOfRowsAbove],
   [averageOfRowsAbove, averageOfRowsAbove, averageOfRowsAbove]
 ];

 var cell = sheet.getRange("B5:D6");
 // This sets the formula to be the sum of the 3 rows above B5.
 cell.setFormulasR1C1(formulas);
  
};