/* File: sheet_to_doc.js 
* Created 4/21/2016
* Small script to grab data from a google sheet 
* and copy into a new report Doc in same folder
*/

function onOpen() {
  // create custom menu to Sheet
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sheet to Doc')
    .addItem('Create Report', 'createReportDoc')
    .addToUi();
}

function createReportDoc() {
  
  // get the data from the spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  // get the data from the spreadsheet
  var header1 = sheet.getRange(4,1,1,3).getValues();
  Logger.log(header1);
  
  var startRow = 5;
  var numColumns = 3;
  
  var data = sheet.getRange(startRow,1,sheet.getLastRow()-4,numColumns).getValues();
  
  // log it - testing purposes only
  Logger.log(data);
  
  // get charts in the spreadsheet
  var charts = sheet.getCharts();
  
  
  // create a new doc in the root folder
  var newDoc = DocumentApp.create('Test Report');
  var newDocId = newDoc.getId();
  var body = newDoc.getBody();
  
  // get the file for this new doc, put into a variable called file
  var file = DriveApp.getFileById(newDocId);
  
  // set the destination folder
  var destFolder = DriveApp.getFolderById('0Bx-btFeE2LPZOXNDa0VXMjM0SWs');
  
  // make a copy of the new Doc in this destination folder
  var timestamp = new Date();
  var newFile = file.makeCopy('Test Report ' + timestamp, destFolder);
  
  var newFileId = newFile.getId();
  
  // write data from the spreadsheet to this new doc
  writeDataToReport(newFileId,data,charts);
  
  // trash the old file
  file.setTrashed(true); 
  
}


function writeDataToReport(id,data,charts) {

  // add the data from Sheet into the Doc
  var doc = DocumentApp.openById(id);  
  var body = doc.getActiveSection();
  
  // Append a document header paragraph.
  var header = body.appendParagraph("Web Report");
  header.setHeading(DocumentApp.ParagraphHeading.HEADING1);

  // Append a section header paragraph.
  var section = body.appendParagraph("Web traffic");
  section.setHeading(DocumentApp.ParagraphHeading.HEADING2);

  // Append a regular paragraph.
  body.appendParagraph("This is a test paragraph");

  // Build a table from the array.
  body.appendTable(data);
  
  // append the charts
  for (var i in charts) {
    body.appendImage(charts[i]);
  }
}