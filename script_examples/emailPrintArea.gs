// menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Email PDF Menu')
      .addItem('Email PDF','emailData')
      .addToUi();
};


// hide rows to control size of print area
function emailData() {
  
  // get sheet id
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var id = ss.getId();
  
  // setup sheets
  var dataSheet = ss.getSheetByName('Sheet1');
  
  var dataURL = ss.getUrl() + "?usp=sharing";
  
  // Ask user how many columns and rows then want to include
  var userChoice = Browser.inputBox('How many rows do you want to print?', 'Please enter a number', Browser.Buttons.OK_CANCEL);
  var rows = parseInt(userChoice) + 2;
  Logger.log(rows);
  
  // find out how many rows in sheet total
  var totalRowCount = dataSheet.getLastRow();
  Logger.log(totalRowCount);
  
  // hide rows
  dataSheet.hideRows(rows, totalRowCount - rows + 2);
  
  // Send the PDF of the spreadsheet to this email address
  // get this from the settings sheet
  var email = 'email@example.com';
  var cc_email = '';
  var bcc_email = '';
 
  // Subject of email message
  var subject = "Print Area testing " + ss.getName() + " - " + new Date().toLocaleString(); 
 
  // Email Body can  be HTML too with your logo image
  var body = "A PDF copy of your data is attached.<br><br>" +
             "To access the live version in Google Sheets, " +
             "<a href='" + dataURL + "'>click this link</a>.";
  
  // Base URL
  var url = "https://docs.google.com/spreadsheets/d/" + id + "/export?";
  
  var url_ext = 'exportFormat=pdf'
      + '&format=pdf'
      + '&size=letter'
      + '&portrait=true'
      + '&fitw=true'
      + '&gid=';
  
  var token = ScriptApp.getOAuthToken();
  var sheetID = dataSheet.getSheetId();
  var sheetName = dataSheet.getName();
  
  var options = 
      {
        headers: {
          'Authorization': 'Bearer ' +  token
        }
        //,"muteHttpExceptions":true
      }
  
  var driveCall = DriveApp.getRootFolder(); // helps initialize first time using the script
  
  // create the pdf
  var response = UrlFetchApp.fetch(url + url_ext + sheetID, options);
  
  // send the email with the PDF attachment
  GmailApp.sendEmail(email, subject, body, {
    cc: cc_email,
    bcc: bcc_email,
    htmlBody: body,
    attachments:[response]     
  });
}