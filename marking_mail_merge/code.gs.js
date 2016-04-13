/* 
* Small script to automatically send emails to all students per selection in marking matrix
* Created by Ben Collins 4/13/16
*/

function sendStudentEmails() {
  // select the range from the Summary sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Summary");
  var lastRow = sheet.getLastRow();
  
  var range = sheet.getRange(4,1,lastRow-3,13).getValues();
  //Logger.log(range.length);
  
  
  // loop over range and send email
  for (var i = 0; i < range.length; i++) {
    if (range[i][11] == "Yes") {
      
      // send email to student by calling sendEmail function
      sendEmail(range[i]);
      
      // create timestamp to add to final column to show when email was sent
      var timestamp = new Date();
      sheet.getRange(i+4,13,1,1).setValue(timestamp);
    };
  }
}

// function to create and send emails
function sendEmail(student) {
  MailApp.sendEmail({
     to: student[1],
     subject: "Feedback for Project",
     htmlBody: 
      "Hi " + student[0] +",<br><br>" +
      "<table  border='1'><tr><td>Section 1 Score</td>" +
      "<td>Section 2 Score</td>" +
      "<td>Section 3 Score</td>" +
      "<td>Overall Score</td></tr>" +
      "<tr><td>" + student[4] + "</td>" +
      "<td>" + student[5] +"</td>" +
      "<td>" + student[6] + "</td>" +
      "<td>" + student[7] + "</td></tr></table>" +
      "<br><b>Positive Notes:</b><br>" +
      student[8] + "<br>" +
      "<br><b>Areas for improvement:</b><br>" +
      student[9] + "<br>" +
      "<br><b>Any other comments:</b><br>" +
      student[10] +
      "<br><br>Marked by: " + student[2]
   });
}