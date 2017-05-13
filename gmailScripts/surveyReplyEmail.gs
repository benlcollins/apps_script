// add menu to Sheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();
   
   ui.createMenu("Send Emails")
     .addItem("Send Email Batch","createEmailNew")
     .addToUi();
 }

/**
 * take the range of data in sheet
 * use it to build an HTML email body
 */
function createEmailNew() {
  var thisWorkbook = SpreadsheetApp.getActiveSpreadsheet();
  var thisSheet = thisWorkbook.getSheetByName('Test Sheet');
  
  // get the data range of the sheet
  var allRange = thisSheet.getDataRange();
  
  // get all the data in this range
  var allData = allRange.getValues();
  
  // get the header row
  var headers = allData.shift();
  
  // create header index map
  var headerIndexes = indexifyHeaders(headers);
  
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
      
      
      var timestamp = sendEmail(row[headerIndexes["Email Address"]],htmlBody);
      thisSheet.getRange(i + 2, headerIndexes["Time replied / Status"] + 1).setValue(timestamp);
    }
    else {
      Logger.log("No email sent for this row");
    }
  });
}
  

/**
 * create index from column headings
 * @param {[object]} headers is an array of column headings
 * @return {{object}} object of column headings as key value pairs with index number
 */
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

/**
 * send email from GmailApp service
 * @param {string} recipient is the email address to send email to
 * @param {string} body is the html body of the email
 * @return {object} new date object to write into spreadsheet to confirm email sent
 */
function sendEmail(recipient,body) {
  
  GmailApp.sendEmail(
    recipient,
    "Thanks for responding about the new course!", 
    "",
    {
      htmlBody: body
    }
  );
  
  return new Date();
}
