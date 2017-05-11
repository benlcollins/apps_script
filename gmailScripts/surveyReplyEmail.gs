function createEmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  var emailAddresses = sheet.getRange(2,2,sheet.getLastRow()-1,1).getValues();
  var info = sheet.getRange(2,11,sheet.getLastRow()-1,1).getValues();
  var whys = sheet.getRange(2,12,sheet.getLastRow()-1,1).getValues();
  var names = sheet.getRange(2,13,sheet.getLastRow()-1,1).getValues();
  var group = sheet.getRange(2,14,sheet.getLastRow()-1,1).getValues();
  var customReply = sheet.getRange(2,15,sheet.getLastRow()-1,1).getValues();
  var status = sheet.getRange(2,16,sheet.getLastRow()-1,1).getValues();
  
  for (var i = 0; i < 8; i++) {
    if (group[i][0] === 1 && !status[i][0]) {
      var   htmlBody = 
          "Hi " + names[i] +",<br><br>" +
            "Thanks for responding and letting me know you're interested in the new course!<br><br>" +
              "<em>Your response:<br><br>" +
                whys[i] + "<br><br>" +
                  info[i] + "</em><br><br>" +
                    customReply[i] + "<br><br>" +
                      "I've had over 200 people respond, which is crazy! So I'll be in touch with a small group about the testing program soon, \n" +
                        "but if you don't hear from me regarding that, I'll still have a special offer for you when the course launches to say thanks!<br><br>" + 
                          "Anything specific you'd like to see in this Data Cleaning and Pivot Table course? Let me know!<br><br>" +
                            "Have a great day.<br><br>" +
                              "Thanks,<br>" +
                                "Ben<br><br>" +
                                  "P.S. I sent you this email directly from my Google Sheet, with a bit of help from Apps Script, using tricks from <a href='http://www.benlcollins.com/spreadsheets/marking-template/'>this tutorial</a>.";
      
      var timestamp = sendEmail(emailAddresses[i],htmlBody);
      sheet.getRange(i+2, 16).setValue(timestamp);
    }
    else if (!status[i][0]){
      sheet.getRange(i+2, 16).setValue("No email sent");
    }
    else {
      // catch all here
    }
  }

}

function sendEmail(recipient,body) {
  Logger.log(recipient[0]);
  
  GmailApp.sendEmail(
    "benlcollins@gmail.com",
    "Thanks for responding about the new course!", 
    "",
    {
      htmlBody: body
    }
  );
  
  return new Date();
}