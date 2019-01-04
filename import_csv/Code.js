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
 * what about xls or xlsx files?
 */

/**
 * add custom menu to run from Google Sheet UI
 */
function onOpen() {
	var ui = SpreadsheetApp.getUi();
	ui.createMenu('Import CSV data')
		//.addItem('Import from file', 'importCSVFromFile')
		.addItem('Import from folder', 'importCSVFromFolder')
		.addItem('Import from Gmail', 'importCSVFromGmail')
		.addToUi();

}

/**
 * ask user if filename contains a header or not
 */
function includesHeader(fileName) {
	var ui = SpreadsheetApp.getUi();
	var response = ui.alert('Does the file ' + fileName + ' have a header row?',ui.ButtonSet.YES_NO);
	return response;
}

/**
 * import CSV data from individual file
 */
function importCSVFromFile(fileName) {
  
	var file = DriveApp.getFilesByName(fileName).next(); 
	var csvData = Utilities.parseCsv(file.getBlob().getDataAsString());

	//Logger.log(csvData);

	var headerRow = includesHeader(fileName);

	if (headerRow == 'YES') { csvData.shift() };

	// paste data into Google Sheet
	var sheet = SpreadsheetApp.getActiveSheet();
	var lastRow = sheet.getLastRow();
	sheet.getRange(lastRow + 1,1,csvData.length, csvData[0].length).setValues(csvData);

}

/**
 * extract csv data from any files in a named Drive folder
 */
function importCSVFromFolder() {

	var folder = DriveApp.getFoldersByName('CSV datasets').next();
	var csvFiles = folder.getFilesByType(MimeType.CSV);

	while (csvFiles.hasNext()) {
		var file = csvFiles.next();
		var fileName = file.getName();

		importCSVFromFile(fileName);

	}

}

/**
 * import csv data from gmail attachments
 */
function importCSVFromGmail() {

	// Example 1: from specific email address and has attachment
	//var threads = GmailApp.search('from: benlcollins@gmail.com has:attachment');

	// Example 1: csv files come from specific email address
	var threads = GmailApp.search('from: benlcollins@gmail.com filename:*.csv');	

	// Example 2: csv files in email with specific subject line
	//var threads = GmailApp.search('subject:CSV Test');

	threads.forEach(function(thread) {
		
		var messageCount = thread.getMessageCount();
		var messages = thread.getMessages();

		messages.forEach(function(message) {
			
			var attachments = message.getAttachments();
			
			attachments.forEach(function(attachment) {
				
				// check if attachment is csv
				if (attachment.getContentType() === 'text/csv') {
					
					var csvData = Utilities.parseCsv(attachment.getDataAsString());

					// paste data into Google Sheet
					var sheet = SpreadsheetApp.getActiveSheet();
					var lastRow = sheet.getLastRow();
					sheet.getRange(lastRow + 1,1,csvData.length, csvData[0].length).setValues(csvData);
				}
			});
		});
	});
}



