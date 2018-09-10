/**
 * Add menu to Sheet
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu("Auto Feedback Tool")
  .addItem("Analyze Feedback","analyzeFeedback")
  .addToUi();
}


/**
 * Get each new row of form data and retrieve the sentiment 
 * scores from the NL API for text in the feedback column.
 */
function analyzeFeedback() {
  
  // get data from the Sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Feedback');
  var allRange = sheet.getDataRange();
  var allData = allRange.getValues();
  
  allData.forEach(function(row,i) {
    if (i !== 0 && row[5] == '') {
      
      Logger.log(i);
      //Logger.log(row);
      //Logger.log("row 4");
      //Logger.log(row[4]);
      
      // call receiveSentiment for each row
      var nlData = retrieveSentiment(row[4]);
      //Logger.log(nlData);
      
      var sentimentScore = nlData.entities[0].sentiment.score;
      var sentimentMagnitude = nlData.entities[0].sentiment.magnitude;
      
      var overallScore = sentimentScore * sentimentMagnitude;
      
      // paste the sentiment to the spreadsheet
      sheet.getRange(i+1, 6, 1, 3).setValues([[sentimentScore,sentimentMagnitude,overallScore]]);
      
      // get the user inputs into variables
      var fullname = row[1];
      var emailAddress = row[2];
      var feedback = row[4];
      
      
      // pass variables to the create draft function
      var timestamp = createDraft(overallScore,fullname,emailAddress,feedback);
      
      //Logger.log(sentimentScore);
      //Logger.log(sentimentMagnitude);
      //Logger.log(overallScore);
      
      sheet.getRange(i+1,9).setValue(timestamp);
    }
      
      
      
      
  });
  
  //Logger.log(allData);
  
}


/**
 * Create a draft email with the feedback and the pre-built message based on sentiment
 * @param
 * @return
 */
function createDraft(overallScore, fullname, emailAddress, feedback) {
  
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
  var   htmlBody = 
        "Hi "+ fullname +",<br><br>" +
          "Thanks for responding to my course feedback questionnaire!<br><br>" +
            response + "<br><br>" +
              "Your feedback:<br><br>" +
                "<i>What did you think of product XYZ?:<br><br>" +
                  feedback +
                    "</i><br><br>" + 
                      "Have a great day!<br><br>" +
                        "Thanks,<br>" +
                          "Ben";
  
  GmailApp.createDraft(
    emailAddress,
    'Thank you for your feedback on the course TEST',
    '',
    {
      htmlBody: htmlBody
    }
  );
  
  var timestamp = new Date(); 
  
  Logger.log(htmlBody);
  
  return timestamp;
}


/**
 * Calls Google Cloud Natural Language API with cell string from my Sheet
 * @param {String} cell The string from a cell in my Sheet
 * @return {Object} the entities and related sentiment present in my string
 */
function retrieveSentiment(cell) {
  
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
  
  //  And make the call
  var response = UrlFetchApp.fetch(apiEndpoint, nlOptions);
  
  return JSON.parse(response);
  
};
