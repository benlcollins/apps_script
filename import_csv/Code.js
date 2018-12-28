/** Program to automatically import csv data from files in google drive folder
 *
 * Step 1: Find the folder (by name) with CSV files
 * Step 2: Extract data from each CSV file
 * Step 3: Paste the data to the Sheet
 * Step 4: Repeat for each file in folder
 * Step 5: Set triggers to automatically look for and import data on a daily/hourly basis
 *
 */

// add custom menu to run from Google Sheet UI
function onOpen() {
	var ui = SpreadsheetApp.getUi();
	ui.createMenu('Import CSV data')
		.addItem('Import from folder', 'importCSVFromFolder')
		.addItem('Test alert', 'includesHeader')
		.addToUi();

}

// main function to control 
function includesHeader() {
	var ui = SpreadsheetApp.getUi();
	var response = ui.alert('Does your data have a header row?',ui.ButtonSet.YES_NO);
	return response;
}


// function to import CSV data
function importCSVFromFolder() {
  
	var file = DriveApp.getFilesByName("testing_csv_2.csv").next();
	var csvData = Utilities.parseCsv(file.getBlob().getDataAsString());

	Logger.log(csvData);

	var headerRow = includesHeader();

	if (headerRow == 'YES') { csvData.shift() };

	var sheet = SpreadsheetApp.getActiveSheet();
	var lastRow = sheet.getLastRow();

	sheet.getRange(lastRow + 1,1,csvData.length, csvData[0].length).setValues(csvData);

}
