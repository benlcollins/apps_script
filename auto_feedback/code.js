/**
 * TO DO:
 * Refactor the spreadsheet data wrangling
 *
 * DONE:
 * Add menu
 * Run the feedback through the NL api to give a sentiment score and magnitude
 * Add sentiment score back into the Sheet
 * Based on the sentiment score, select a boilerplate response tone (positive to negative)
 * Create a draft email in Gmail with boilerplate email text and the original feedback
 * Apps script detects when email has been sent and to which address, and marks it as sent in my Sheet
 * Once a day apps script checks how many new drafts have been created and sends me a reminder, "You have 3 course feedback drafts waiting for your attention"
 * Add error handling so that if the cloud natural language api craps out, it assumes a score of 0 and creates the draft email that way
 * Setup triggers to run draft alert and confirm emails sent on a daily basis
 *
 */


/**
 * Add menu to Sheet
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu("Auto Feedback Tool")
  .addItem("Analyze Feedback","analyzeFeedback")
  .addItem("Confirm Sent","confirmEmailSent")
  .addItem("Drafts Waiting?","draftsWaitingAlert")
  .addToUi();
}


/**
 * Get each new row of form data and retrieve the sentiment 
 * scores from the NL API for text in the feedback column.
 */
function analyzeFeedback() {
  
  // get data from the Sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Form Responses 1');
  var allRange = sheet.getDataRange();
  var allData = allRange.getValues();
  
  // extract the header row
  var headers = allData.shift();
  Logger.log(headers);
  
  
  allData.forEach(function(row,i) {
    if (row[8] == '') {
      
      //Logger.log(i);
      //Logger.log(row[6]);

      // get the user inputs into variables
      
      var emailAddress = row[1];
      var courseName = row[2];
      var feedbackSource = row[3];
      var feedbackWhy = row[4];
      var feedbackHope = row[5];
      var feedbackElse = row[6];
      var fullname = row[7];
      
      
      // call receiveSentiment for each row
      if (row[6] != '') {
        
        var nlData = retrieveSentiment(feedbackElse);

        Logger.log(nlData);

        // set to zero if no sentiment score is returned from the api
        var sentimentScore = nlData.entities[0] ? nlData.entities[0].sentiment.score : 0;
        var sentimentMagnitude = nlData.entities[0] ? nlData.entities[0].sentiment.magnitude : 0;

      }
      else {
        
        // set to zero if the comment cell is blank in Sheet
        var sentimentScore = 0;
        var sentimentMagnitude = 0;

      }
      
      // create an overall score
      var overallScore = sentimentScore * sentimentMagnitude;
      
      // paste the sentiment to the spreadsheet
      sheet.getRange(i+2, 11, 1, 3).setValues([[sentimentScore,sentimentMagnitude,overallScore]]);
      
      // pass variables to the create draft function
      createDraft(overallScore,fullname,emailAddress,courseName,feedbackWhy,feedbackHope,feedbackElse);

      var d = new Date();
      
      sheet.getRange(i+2,9).setValue(d);
      
    }
      
      
      
      
  });
  
}


/**
 * Create a draft email with the feedback and the pre-built message based on sentiment
 * @param
 * @return
 */
function createDraft(overallScore,fullname,emailAddress,courseName,feedbackWhy,feedbackHope,feedbackElse) {
  
  // get pre-populated data from the Sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Email Templates');
  
  if (overallScore > 0) {
    // positive sentiment
    var response = sheet.getRange(2,2).getValue();
  }
  else if (overallScore == 0) {
    // neutral sentiment
    var response = sheet.getRange(3,2).getValue();
  }
  else {
    // negative sentiment
    var response = sheet.getRange(4,2).getValue();
  }
  
  // create the draft email
  var subjectLine = 'Thanks for your feedback on the Google Sheets courses';
  
  var htmlBody = 
        "Hi "+ fullname +",<br><br>" +
          "Thanks for responding to my course feedback questionnaire!<br><br>" +
            response + "<br><br>" +
              "Your feedback:<br><br>" +
                "<i>Why are you taking the course(s)?<br><br>" +
                  feedbackWhy + "<br><br>" +
                    "What are you hoping to get out of the course(s)?<br><br>" +
                      feedbackHope + "</i><br><br>" + 
                          "Have a great day!<br><br>" +
                            "Thanks,<br>" +
                              "Ben";
  
  GmailApp.createDraft(
    emailAddress,
    subjectLine,
    '',
    {
      htmlBody: htmlBody
    }
  );
  
  Logger.log(htmlBody);
}


/**
 * Confirm email sent
 * run this once a day maybe?
 * @param
 * @return
 */
function confirmEmailSent() {
  
  // move this to a global variable or pass it in
  var subjectLine = 'Thanks for your feedback on the Google Sheets courses';
  
  // find email in Sent emails folder that matches the email address and subject line
  // check all the rows in my dataset with a draft date next to them
  // add a new timestamp in the sent column once sending has been confirmed
  
  // get data from the Sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Form Responses 1');
  var allRange = sheet.getDataRange();
  var allData = allRange.getValues(); 

  allData.shift(); // remove header row

  //Logger.log(allData);

  allData.forEach(function(row,i) {

    if (row[9] == '') {
      
      Logger.log(i);
      Logger.log(row);
      var emailAddress = row[1];
      Logger.log(emailAddress);

      if (GmailApp.search("in:sent to:" + emailAddress + " subject:" + subjectLine)[0]) {
        
        Logger.log("result!");

        var d = new Date();
        sheet.getRange(i+2, 10).setValue(d);
      }

    }
  }); 
  
}



/**
 * Alert me that draft emails are waiting for me to review and send
 * run this once a day by trigger
 * @param
 * @return
 */
function draftsWaitingAlert() {
  
  // get data from the Sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Form Responses 1');
  var allRange = sheet.getDataRange();
  var allData = allRange.getValues();
  allData.shift(); // remove header row
  
  var draftCount = 0;
  
  allData.forEach(function(row,i) {
    if (row[9] == '') {
      draftCount++;
    }
  });
  
  if (draftCount > 0) {
    GmailApp.sendEmail(
      "ben@benlcollins.com", 
      "Draft emails waiting re Course Feedback", 
      "You have " + draftCount + " draft emails relating to course feedback waiting for review."
    );
  }
  
}



/**
 * Calls Google Cloud Natural Language API with cell string from my Sheet
 * @param {String} cell The string from a cell in my Sheet
 * @return {Object} the entities and related sentiment present in my string
 */
function retrieveSentiment(cell) {
  
  //Logger.log(cell);
  
  var apiEndpoint = 'https://language.googleapis.com/v1/documents:analyzeEntitySentiment?key=' + apiKey;
  
  // Create our json request, w/ text, language, type & encoding
  var nlData = {
    document: {
      language: 'en-us',
      type: 'PLAIN_TEXT',
      content: cell
    },
    encodingType: 'UTF8'
  };
  
  //  Package all of the options and the data together for the call
  var nlOptions = {
    method : 'post',
    contentType: 'application/json',  
    payload : JSON.stringify(nlData)
  };
  
  //  Try fetching the natural language api
  try {
    
    // return the parsed JSON data if successful
    var response = UrlFetchApp.fetch(apiEndpoint, nlOptions);
    return JSON.parse(response);
    
  } catch(e) {
    
    // log the error message and return null if not successful
    Logger.log("Error fetching the Natural Language API: " + e);
    return null;
  }
  
  
  
};

