/* 
* Small script to automatically send emails and slack messages to all students per selection in marking matrix
* Created by Ben Collins 4/13/16
*/

function sendStudentScores() {
  // select the range from the Summary sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Summary");
  var lastRow = sheet.getLastRow();
  
  var range = sheet.getRange(4,1,lastRow-3,15).getValues();
  //Logger.log(range.length);
  
  // create timestamp to mark when communication was sent
  var timestamp = new Date();
  
  // loop over range and send communication if "Yes" option chosen
  for (var i = 0; i < range.length; i++) {
    if (range[i][12] == "Yes") {
      
      // choose email, slack or both channels
      switch (range[i][13]) {
        case "Email":
          // send email to student by calling sendEmail function
          sendEmail(range[i]);
          break;
        
        case "Slack":
          // post message to slack
          sendToSlack(range[i]);
          break;
          
        case "Both":
          // send email and post to Slack
          sendEmail(range[i]);
          sendToSlack(range[i]);
      }
      
      // add timestamp to final column to show when communication was sent
      sheet.getRange(i+4,15,1,1).setValue(timestamp);
    };
  }
}

// function to create and send emails
function sendEmail(student) {
  var timestamp = new Date();
  MailApp.sendEmail({
     to: student[1],
     subject: "Feedback for Project",
     htmlBody: 
      "Hi " + student[0] +",<br><br>" +
      "Here are your project scores and feedack. Let us know if you have any questions!<br><br>" +
      "<table  border='1'><tr><td>Section 1 Score</td>" +
      "<td>Section 2 Score</td>" +
      "<td>Section 3 Score</td>" +
      "<td>Overall Score</td></tr>" +
      "<tr><td>" + student[5] + "</td>" +
      "<td>" + student[6] +"</td>" +
      "<td>" + student[7] + "</td>" +
      "<td>" + student[8] + "</td></tr></table>" +
      "<br><b>Positive Notes:</b><br>" +
      student[9] + "<br>" +
      "<br><b>Areas for improvement:</b><br>" +
      student[10] + "<br>" +
      "<br><b>Any other comments:</b><br>" +
      student[11] +
      "<br><br>Marked by: " + student[3] +
      "<br>Date: " + timestamp +
      "<br><br>Sent care of Marking Mail Merge tool built by <a href='http://www.benlcollins.com/'>Ben Collins</a>"
   });
}


// function to send message to Slack
function sendToSlack(student) {
  var timestamp = new Date();
  
  // custom slack webhook
  // change the XXXXX's to your own slack webhook. Get it from: 
  // https://my.slack.com/services/new/incoming-webhook/
  var url = "https://hooks.slack.com/services/XXXXXXXXX/XXXXXXXXX/XXXXXXXXXXXXXXXXXXXXXXXX";
  
  var payload = {
    "channel": "@"+student[2],
    "username": "ben",
    "text": "Hi " + student[0] +
      "\n Here are your project scores and feedack. Let us know if you have any questions! \n" +
      "\n Section 1 Score: " + student[5] +
      "\n Section 2 Score: " + student[6] +
      "\n Section 3 Score: " + student[7] +
      "\n Overall Score: " + student[8] +
      "\n *Positive notes:* " + student[9] +
      "\n *Areas for improvement:* " + student[10] +
      "\n *Any other comments:* " + student[11] +
      "\n \n Marked by: " + student[3] +
      "\n Date: " + timestamp +
      "\n \n Sent care of Marking Mail Merge tool built by <http://www.benlcollins.com/|benlcollins>",
    "icon_emoji": ":inbox_tray:"
  };

  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };
  
  return UrlFetchApp.fetch(url,options);
}
