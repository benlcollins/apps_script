// for publishing as web app
function doGet(e) {
  if (!e.parameter.page) {
    // default for when no specific page requested
    return HtmlService.createTemplateFromFile('index.html')
      .evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  }
  else {
    var t = HtmlService.createTemplateFromFile(e.parameter.page);
    
    // pushing variable to template
    // more details here:
    // https://developers.google.com/apps-script/guides/html/templates#calling_apps_script_apis_directly
    t.data = e.parameter.value;
    
    return t.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
  }
}

// submit form
function submitForm(form) {
  
  // get the form data
  var data = form.memberName;
  
  // select addresses (make this auto-populate based on form entries)
  var adminEmail = "xyz@xyz.com";
  
  // subject and time of email
  var subject = "Report " + (new Date()).toString();
  
  // body of email
  // includes link
  // want this link to open up the other HTML page (step 1)
  // and pre-populate with the relevant data from first form (step 2)
  // Replace XXXXXXX with correct url for form
  var htmlBody = "Email sent on " + (new Date()).toString() + "<br><br>" +
                 "Submitted by: " + data + "<br><br>" +
                 "<a href='https://script.google.com/macros/s/XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX/dev?page=supplierForm&value=" 
                 + data + "'>Respond to RFP</a>";
  
  Logger.log(htmlBody);
  
  // send the email
  MailApp.sendEmail(adminEmail, subject, htmlBody,{htmlBody:htmlBody});
  
  return "Form submitted successfully";
}