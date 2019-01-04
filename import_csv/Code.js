/** Program to automatically import csv data from files in google drive folder
 *
 * Step 1: Find the folder (by name) with CSV files
 * Step 2: Extract data from each CSV file
 * Step 3: Paste the data to the Sheet
 * Step 4: Repeat for each file in folder
 * Step 5: Automatically identify header rows in datasets
 * Step 6: Set triggers to automatically look for and import data on a daily/hourly basis
 *
 * Other considerations:
 * what if there are two files with same name?
 */

// add custom menu to run from Google Sheet UI
function onOpen() {
	var ui = SpreadsheetApp.getUi();
	ui.createMenu('Import CSV data')
		//.addItem('Import from file', 'importCSVFromFile')
		.addItem('Import from folder', 'importCSVFromFolder')
		.addToUi();

}

// main function to control 
function includesHeader(fileName) {
	var ui = SpreadsheetApp.getUi();
	var response = ui.alert('Does the file ' + fileName + ' have a header row?',ui.ButtonSet.YES_NO);
	return response;
}


// function to import CSV data from file
function importCSVFromFile(fileName) {
  
	var file = DriveApp.getFilesByName(fileName).next(); 
	var csvData = Utilities.parseCsv(file.getBlob().getDataAsString());

	Logger.log(csvData);

	var headerRow = includesHeader(fileName);

	if (headerRow == 'YES') { csvData.shift() };

	var sheet = SpreadsheetApp.getActiveSheet();
	var lastRow = sheet.getLastRow();

	sheet.getRange(lastRow + 1,1,csvData.length, csvData[0].length).setValues(csvData);

}


// generalized to extract csv data from any files in a Drive folder
function importCSVFromFolder() {

	var folder = DriveApp.getFoldersByName('CSV datasets').next();
	var csvFiles = folder.getFilesByType(MimeType.CSV);

	while (csvFiles.hasNext()) {
		var file = csvFiles.next();
		var fileName = file.getName();

		importCSVFromFile(fileName);

	}

}





