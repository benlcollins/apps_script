/*
 * Script to extract all the email addresses for a particular folder label in gmail
 *
 */


// add menu to Sheet
function onOpen() {
  
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Extract Emails')
      .addItem('Extract Emails...', 'extractEmails')
      .addItem('Extract Emails + Subjects...','extractEmailsSubjects')
      .addToUi();
}


// extract emails from label in Gmail
function extractEmails() {
  
  // get the spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var label = sheet.getRange(1,2).getValue();
  
  Logger.log(label);
  
  // get all email threads that match label from Sheet
  var threads = GmailApp.search ("label:" + label);
  
  //Logger.log(threads.length); // 210 which matches my count in gmail
  
  // get all the messages for the current batch of threads
  var messages = GmailApp.getMessagesForThreads (threads);
  
  /*
   * old way of looping
   *
  for (var i = 0 ; i < messages.length; i++) {
   for (var j = 0; j < messages[i].length; j++) {
     Logger.log("from: " + messages[i][j].getFrom());
   }
  }
  */
  
  var emailArray = [];
  
  // get array of email addresses
  messages.forEach(function(message) {
    //fromEmails.push(message.getFrom());
    //Logger.log(message[0].getFrom());
    message.forEach(function(d) {
      emailArray.push(d.getFrom(),d.getTo());
      //Logger.log(d.getFrom());
      //Logger.log(d.getReplyTo());
      //Logger.log(d.getTo());
    });
  });
  
  //Logger.log(emailArray.length);
  //Logger.log(emailArray);
  
  // de-duplicate the array
  var uniqueEmailArray = emailArray.filter(function(item, pos) {
    return emailArray.indexOf(item) == pos;
  });
  
  //Logger.log(uniqueEmailArray.length);
  
  var cleanedEmailArray = uniqueEmailArray.map(function(el) {
    var name = "";
    var email = "";
    
    var matches = el.match(/\s*"?([^"]*)"?\s+<(.+)>/);
    
    if (matches) {
      //Logger.log(matches);
      name = matches[1]; 
      email = matches[2];
    }
    else {
      name = "N/k";
      email = el;
    }
    
    return [name,email];
  }).filter(function(d) {
    if (
         d[1] !== "benlcollins@gmail.com" &&
         d[1] !== "drive-shares-noreply@google.com" &&
         d[1] !== "wordpress@www.benlcollins.com"
       ) {
      return d;
    }
  });
  
  //Logger.log(cleanedEmailArray);
  
  // clear any old data
  sheet.getRange(4,1,sheet.getLastRow(),2).clearContent();
  
  // paste in new names and emails and sort by email address A - Z
  sheet.getRange(4,1,cleanedEmailArray.length,2).setValues(cleanedEmailArray).sort(2);
 
}

// extract emails from label in Gmail
function extractEmailsSubjects() {
  
  // get the spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var label = sheet.getRange(1,2).getValue();
  
  Logger.log(label);
  
  // get all email threads that match label from Sheet
  var threads = GmailApp.search("label:" + label);
  
  //Logger.log(threads.length); // 210 which matches my count in gmail
  
  // get all the messages for the current batch of threads
  var messages = GmailApp.getMessagesForThreads (threads);
  
  /*
   * old way of looping
   *
  for (var i = 0 ; i < messages.length; i++) {
   for (var j = 0; j < messages[i].length; j++) {
     Logger.log("from: " + messages[i][j].getFrom());
   }
  }
  */
  
  var emailArray = [];
  
  // get array of email addresses
  messages.forEach(function(message) {
    //fromEmails.push(message.getFrom());
    //Logger.log(message[0].getFrom());
    message.forEach(function(d) {
      emailArray.push([d.getFrom(),d.getTo(),d.getSubject()]);
      //Logger.log(d.getFrom());
      //Logger.log(d.getReplyTo());
      //Logger.log(d.getTo());
    });
  });
  
  Logger.log(emailArray.length);
  Logger.log(emailArray);
  
  
  //Logger.log(cleanedEmailArray);
  
  // clear any old data
  sheet.getRange(4,1,sheet.getLastRow(),3).clearContent();
  
  // paste in new names and emails and sort by email address A - Z
  sheet.getRange(4,1,emailArray.length,3).setValues(emailArray).sort(2);
  
}