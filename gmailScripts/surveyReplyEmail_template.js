// add menu to Sheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu("Send Emails")
  .addItem("Send Email Batch","createEmail")
  .addToUi();
}

/**
 * take the range of data in sheet
 * use it to build an HTML email body
 */
function createEmail() {
  var thisWorkbook = SpreadsheetApp.getActiveSpreadsheet();
  var thisSheet = thisWorkbook.getSheetByName('Form Responses 1');

  // get the data range of the sheet
  var allRange = thisSheet.getDataRange();
  
  // get all the data in this range
  var allData = allRange.getValues();
  
  // get the header row
  var headers = allData.shift();
  
  // create header index map
  var headerIndexes = indexifyHeaders(headers);
  
  allData.forEach(function(row,i) {
    if (!row[headerIndexes["Status"]]) {
      var   htmlBody = 
        "Hi " + row[headerIndexes["What is your name?"]] +",<br><br>" +
          "Thanks for responding to my questionnaire!<br><br>" +
            "<em>Your choice:<br><br>" +
              row[headerIndexes["Choose a number between 1 and 5?"]] + "</em><br><br>" +
                row[headerIndexes["Custom Reply"]] + "<br><br>" + 
                  "Have a great day.<br><br>" +
                    "Thanks,<br>" +
                      "Ben";
      
      var timestamp = sendEmail(row[headerIndexes["Email Address"]],htmlBody);
      thisSheet.getRange(i + 2, headerIndexes["Status"] + 1).setValue(timestamp);
    }
    else {
      Logger.log("No email sent for this row: " + i + 1);
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
    "Thanks for responding to the questionnaire!", 
    "",
    {
      htmlBody: body
    }
  );
  
  return new Date();
}