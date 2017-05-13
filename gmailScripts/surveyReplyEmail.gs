function onOpen() {
  var ui = SpreadsheetApp.getUi();
   
   ui.createMenu("Send Emails")
     .addItem("Send Email Batch","createEmailNew")
     .addToUi();
 }

// change the range references to work with column heading names, instead of absolute references so they don't break when I move them around

function createEmailNew() {
  var thisWorkbook = SpreadsheetApp.getActiveSpreadsheet();
  var thisSheet = thisWorkbook.getActiveSheet();
  
  // get the data range of the sheet
  var allRange = thisSheet.getDataRange();
  
  // get all the data in this range
  var allData = allRange.getValues();
  
  // get the header row
  var headers = allData.shift();
  //Logger.log(headers);
  
  // create header index map
  var headerIndexes = indexifyHeaders(headers);
  //Logger.log(JSON.stringify(headerIndexes));
  
  /*
  // test to see if header index works
  allData.forEach(function(row) {
    //Logger.log(row[headerIndexes["What's your name?"]]);
  });
  
  var emailAddresses = [];
  allData.forEach(function(row) {
    emailAddresses.push(row[headerIndexes["Email Address"]]);
  });
  Logger.log(emailAddresses);
  */
  
  allData.forEach(function(row,i) {
    if (row[headerIndexes["Reply Group"]] === 1 && !row[headerIndexes["Time replied / Status"]]) {
      var   htmlBody = 
          "Hi " + row[headerIndexes["What's your name?"]] +",<br><br>" +
            "Thanks for responding and letting me know you're interested in the new course!<br><br>" +
              "<em>Your response:<br><br>" +
                row[headerIndexes["Why do you want to take this course?"]] + "<br><br>" +
                  row[headerIndexes["Any other information you'd like to share with me? I'm happy to hear from you..."]] + "</em><br><br>" +
                    row[headerIndexes["My Reply"]] + "<br><br>" +
                      "I've had over 250 people respond, which is crazy! So I'll be in touch with a small group about the testing program soon, \n" +
                        "but if you don't hear from me regarding that, I'll still have a special offer for you when the course launches to say thanks!<br><br>" + 
                          "Anything specific you'd like to see in this Data Cleaning and Pivot Table course? Let me know!<br><br>" +
                            "Have a great day.<br><br>" +
                              "Thanks,<br>" +
                                "Ben<br><br>" +
                                  "P.S. I sent you this email directly from my Google Sheet, with a bit of help from Apps Script, using tricks from <a href='http://www.benlcollins.com/spreadsheets/marking-template/'>this tutorial</a>.";
      
      Logger.log(htmlBody);
    
    }
    else {
      //Logger.log("No");
    }
  });
                  
    
    /*
    if (row[headerIndexes["What's your name?"]][i][0] === 1 && !row[headerIndexes["What's your name?"]][i][0]) {
      var   htmlBody = 
          "Hi " + row[headerIndexes["What's your name?"]][i] +",<br><br>" +
            "Thanks for responding and letting me know you're interested in the new course!<br><br>" +
              "<em>Your response:<br><br>" +
                //whys[i] + "<br><br>" +
                  //info[i] + "</em><br><br>" +
                    //customReply[i] + "<br><br>" +
                      "I've had over 250 people respond, which is crazy! So I'll be in touch with a small group about the testing program soon, \n" +
                        "but if you don't hear from me regarding that, I'll still have a special offer for you when the course launches to say thanks!<br><br>" + 
                          "Anything specific you'd like to see in this Data Cleaning and Pivot Table course? Let me know!<br><br>" +
                            "Have a great day.<br><br>" +
                              "Thanks,<br>" +
                                "Ben<br><br>" +
                                  "P.S. I sent you this email directly from my Google Sheet, with a bit of help from Apps Script, using tricks from <a href='http://www.benlcollins.com/spreadsheets/marking-template/'>this tutorial</a>.";
      
      //var timestamp = sendEmail(emailAddresses[i],htmlBody); 
      var timestamp = newDate();
      //sheet.getRange(i+2, 20).setValue(timestamp);
    }
    else if (!status[i][0]){
      //sheet.getRange(i+2, 20).setValue("No email sent");
    }
    else {
      // catch all here
    }
    */
  
}
  

// create index from column headings
function indexifyHeaders(headers) {
  
  var index = 0;
  return headers.reduce (
    // callback function
    function(p,c) {
    
      //skip cols with blank headers
      if (c) {
        // can't have duplicate column names
        if (p.hasOwnProperty(c)) {
          throw new Error('duplicate column name: ' + c);
        }
        p[c] = index;
      }
      index++;
      return p;
    },
    {} // initial value for reduce function to use as first argument
  );
}




function createEmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  var emailAddresses = sheet.getRange(2,2,sheet.getLastRow()-1,1).getValues();
  var info = sheet.getRange(2,11,sheet.getLastRow()-1,1).getValues();
  var whys = sheet.getRange(2,12,sheet.getLastRow()-1,1).getValues();
  var names = sheet.getRange(2,13,sheet.getLastRow()-1,1).getValues();
  var group = sheet.getRange(2,18,sheet.getLastRow()-1,1).getValues();
  var customReply = sheet.getRange(2,19,sheet.getLastRow()-1,1).getValues();
  var status = sheet.getRange(2,20,sheet.getLastRow()-1,1).getValues();
  
  for (var i = 0; i < emailAddresses.length; i++) {
    if (group[i][0] === 1 && !status[i][0]) {
      var   htmlBody = 
          "Hi " + names[i] +",<br><br>" +
            "Thanks for responding and letting me know you're interested in the new course!<br><br>" +
              "<em>Your response:<br><br>" +
                whys[i] + "<br><br>" +
                  info[i] + "</em><br><br>" +
                    customReply[i] + "<br><br>" +
                      "I've had over 250 people respond, which is crazy! So I'll be in touch with a small group about the testing program soon, \n" +
                        "but if you don't hear from me regarding that, I'll still have a special offer for you when the course launches to say thanks!<br><br>" + 
                          "Anything specific you'd like to see in this Data Cleaning and Pivot Table course? Let me know!<br><br>" +
                            "Have a great day.<br><br>" +
                              "Thanks,<br>" +
                                "Ben<br><br>" +
                                  "P.S. I sent you this email directly from my Google Sheet, with a bit of help from Apps Script, using tricks from <a href='http://www.benlcollins.com/spreadsheets/marking-template/'>this tutorial</a>.";
      
      var timestamp = sendEmail(emailAddresses[i],htmlBody);      
      sheet.getRange(i+2, 20).setValue(timestamp);
    }
    else if (!status[i][0]){
      sheet.getRange(i+2, 20).setValue("No email sent");
    }
    else {
      // catch all here
    }
  }

}

function sendEmail(recipient,body) {
  Logger.log(recipient[0]);
  
  GmailApp.sendEmail(
    //recipient[0],
    "benlcollins@gmail.com", // use for testing
    "Thanks for responding about the new course!", 
    "",
    {
      htmlBody: body
    }
  );
  
  return new Date();
}