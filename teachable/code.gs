/** 
* Teachable Course Certificate Tool
* Built by Ben Collins, 2018
* https://www.benlcollins.com/apps-script/teachable-certificates
*/


/**
* @description Constants
*/
var TEMPLATE_URL = "https://docs.google.com/document/d/{doc_id}/edit";
var CERTIFICATE_TITLE = "Ben Collins Data School Student Certificate";


/**
* onOpen function to create custom menu in Sheet
*/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Certificates')
      .addItem('Send Certificates', 'sendCertificates')
      .addToUi(); 
}


/**
* doPost(e) function to listen for incoming post requests and display data in Sheet
* @param {e} post request - the post request
* @returns null
*/
function doPost(e) {
  if (typeof e !== 'undefined') { 

    // extract the relevant data
    var data = JSON.parse(e.postData.contents);
    var completion_date = data.created;
    var user_name = data.object.user.name;
    var course_name = data.object.course.name;
    var user_email = data.object.user.email;
    
    // put into array for Sheet
    var row = [];
    
    row.push(
      completion_date,
      user_name,
      course_name,
      user_email
    );
    
    // setup the Sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Sheet1');
    
    // paste data into new row of Sheet
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1,1,1,4).setValues([row]);
  }
  return null;
}




/**
* sendCertificates function to take data from Sheet, create and send certificates
*/
function sendCertificates() {
  
  // grab the data from the spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Sheet1');
  
  // count completed certificates from final column
  var completedCount = sheet.getRange("E1:E").getValues().filter(String).length;
  
  // get pending rows of data
  var pendingData = sheet.getRange(completedCount + 1, 1, sheet.getLastRow() - completedCount, 5).getValues();

  pendingData.forEach(function(row, i) {
    try {
      // get the student variables
      var teachable_date = row[0];
      var name = row[1];
      var course = row[2];
      var email = row[3]; 
      
      // fix date so it's better format
      var year = teachable_date.substring(0,4);
      var month = teachable_date.substring(5,7) * 1;
      var months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
      var day = teachable_date.substring(8,10);
      var formatted_date = day + " " + months[month - 1] + " " + year;
      
      // create certificate
      var cert = createCertificate(name, course, formatted_date); 
      var doc = cert[0]; 
      
      // turn certificate into pdf
      var docblob = doc.getAs('application/pdf');
      docblob.setName(doc.getName() + ".pdf");
      var file = DriveApp.createFile(docblob);
      
      // email certificate
      emailCertificate(name, course, email, file);
      
      // trash the doc and the file
      var doc_file = DriveApp.getFileById(cert[1]); 
      doc_file.setTrashed(true);
      file.setTrashed(true);
      
      // mark this row as complete in the dataset
      sheet.getRange(completedCount + 1 + i,5).setValue(formatted_date);
    }
    catch(e) {
      Logger.log(e.message);
    }
  });
}


/**
* createCertificate function to create the certificate from the Google Doc template
* @param {name} string - student name
* @param {course} string - course name
* @param {date} string - course completion date
* @returns {array} certificate and certificate ID
*/
function createCertificate(name, course, date) {
  
  // extract the Google Doc template ID from the URL
  var rx = /\/d\/(.*)\/edit/i;
  var templateId = TEMPLATE_URL.match(rx)[1];

  // make copy of the template and assign to variable doc
  var certificate = DriveApp.getFileById(templateId).makeCopy(CERTIFICATE_TITLE);
  var certificateID = certificate.getId();
  var doc = DocumentApp.openById(certificateID);
  
  // update the certificate
  var doc = updateCertificate(doc, name, course, date);
  
  return [doc, certificateID];
}


/**
* updateCertificate function to update the certificate with student data
* @param {doc} doc - the certificate being created
* @param {name} string - student name
* @param {course} string - course name
* @param {date} string - course completion date
* @returns {doc} file - the updated certificate
*/
function updateCertificate(doc, name, course, date) {

  var body = doc.getBody();
  
  // find the {{variables}} in the google doc and replace with the actuals  
  body.replaceText("{{student_name}}", name);
  body.replaceText("{{course_name}}", course);
  body.replaceText("{{date}}", date);
  doc.saveAndClose();
  
  return doc;
}


/**
* emailCertificate function to email the certificate file to the student
* @param {name} string - student name
* @param {course} string - course name
* @param {email} string - student email
* @param {file} file - pdf copy of the certificate
*/
function emailCertificate(name, course, email, file) {

  var htmlBody = 'Hi ' + name + ',<br><br>' +
                       'Congratulations on recently completing the <strong>' + course + '</strong>!<br><br>' +
                       'This is a fantastic achievement and I want to be the first to congratulate you. Here\'s your certificate to acknowledge your hard work and successful course completion.<br><br>' +
                       'Best wishes for your continued success and I hope to see you again soon at <a href="https://www.benlcollins.com/">benlcollins.com</a>.<br><br>' +
                       'Sincerely,<br>' +
                       'Ben';
  
  if (email) {
    GmailApp.sendEmail(
      email, 
      'Congratulations from Ben Collins Data School!',
      '',
      {
        htmlBody: htmlBody,
        attachments: file,
        name: 'Congratulations!'
      }
    );
  }
}