// ------------------------------------------------------------------------------------------
// Dashboard Code
// Ben Collins
// Learn more at: http://www.benlcollins.com/
// ------------------------------------------------------------------------------------------


// ------------------------------------------------------------------------------------------
// Menu in Google Sheet
// ------------------------------------------------------------------------------------------

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Dashboard Menu')
      .addItem('Email Dashboard','emailDashboard')
      .addItem('Save Social Data','saveSocialData')
      .addItem('Save Alexa Data','saveAlexaData')
      .addToUi();
};



// ------------------------------------------------------------------------------------------
// Convert dashboard to PDF and email a copy to user
// ------------------------------------------------------------------------------------------

function emailDashboard() {
  
  // setup sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboardSheet = ss.getSheetByName('Dashboard');
  var settingsSheet = ss.getSheetByName('settings');
  
  // Send the PDF of the spreadsheet to this email address
  // get this from the settings sheet
  var email = settingsSheet.getRange(6,2).getValue();
 
  // Subject of email message
  var subject = "Dashboard PDF generated from " + ss.getName(); 
 
  // Email Body can  be HTML too with your logo image
  var body = "Here is your Dashboard.";
  
  // Base URL
  var url = "https://docs.google.com/spreadsheets/d/1uAZIowho5IkmN0lFVxg-XR9T7fdg1J45P44KQ92UKME/export?";
  
  var url_ext = 'exportFormat=pdf&format=pdf'        // export as pdf / csv / xls / xlsx
  + '&size=letter'                       // paper size legal / letter / A4
  + '&portrait=true'                    // orientation, false for landscape
  + '&fitw=true'                         // fit to page width, false for actual size   // &source=labnol
  + '&sheetnames=false&printtitle=false' // hide optional headers and footers
  + '&pagenumbers=false&gridlines=false' // hide page numbers and gridlines
  + '&fzr=false'                         // do not repeat row headers (frozen rows) on each page
  + '&gid=';                             // the sheet's Id
  
  var token = ScriptApp.getOAuthToken();
  var sheetID = dashboardSheet.getSheetId();
  var sheetName = dashboardSheet.getName();
  
  var options = 
      {
        headers: {
          'Authorization': 'Bearer ' +  token
        }
        //,"muteHttpExceptions":true
      }
  
  //var bogus = DriveApp.getRootFolder();
  
  var response = UrlFetchApp.fetch(url + url_ext + sheetID, options);
  
  // If allowed to send emails, send the email with the PDF attachment
  if (MailApp.getRemainingDailyQuota() > 0) 
    Logger.log("set to go");
    GmailApp.sendEmail(email, subject, body, {
      htmlBody: body
      ,attachments:[response]     
    });
  
}