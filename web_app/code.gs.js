// for publishing as web app
function doGet(e) {
  return HtmlService.createTemplateFromFile('index.html')
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}


// submit RFP form
function submitRFP(form) {
  Logger.log("RFP submitted successfully!");
  
  // setup lock service
  var lock = LockService.getPublicLock();
  lock.waitLock(30000);  // wait 30 seconds before conceding defeat.
  
  try {
    
    /* recording content to spreadsheet section */
    // set where we write the data - you could write to multiple/alternate destinations
    var doc = SpreadsheetApp.openById('1kL5RSCuLtpgLDrMWDT5WMWFKVPceTHiNs7EW0ljq_Z4');
    var sheet = doc.getSheetByName('Results');
    var RFPsheet = doc.getSheetByName('RFPasPDF');
    
    var timestamp = new Date();
    
    var data = [[timestamp,
                 form.memberName,
                 form.memberPOC,
                 form.memberEmail,
                 form.memberCell,
                 form.categories,
                 form.suppliers,
                 form.suppliersPOC,
                 form.supplierNonRD,
                 form.clientName,
                 form.brandName,
                 form.programName,
                 form.programDesc,
                 form.proposalDueDate,
                 form.programStartDate,
                 form.programEndDate,
                 form.totalBudget,
                 form.itemName1,
                 form.itemDesc1,
                 form.itemQuantity1,
                 form.itemCost1,
                 form.itemName2,
                 form.itemDesc2,
                 form.itemQuantity2,
                 form.itemCost2
                ]];
    
    var outputRange = sheet.getRange(sheet.getLastRow()+1, 1, 1, 25);
    outputRange.setValues(data.map(function(d){
      return d; 
    }));
    
    // formula to multiply the item 1 quantity and item description and add into sheet
    sheet.getRange(sheet.getLastRow(), 26).setFormula('=T'+sheet.getLastRow()+'*U'+sheet.getLastRow());
    
    // formula to work out total cost of item 2
    sheet.getRange(sheet.getLastRow(), 28).setFormula('=X'+sheet.getLastRow()+'*Y'+sheet.getLastRow());

    
    /* uploading file section */
    var dropbox = "Test Images XYZ";
    var folder, folders = DriveApp.getFoldersByName(dropbox);
    
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(dropbox);
    }
    
    // file 1 upload
    var blob1 = form.updloadFile1;    
    var file1 = folder.createFile(blob1);    
    file1.setDescription("Uploaded by " + form.memberName);
    var file1Url = file1.getUrl();
    //Logger.log(file1Url);
    sheet.getRange(sheet.getLastRow(), 27).setFormula('=hyperlink("' + file1Url + '","Image")');  // e.g. =HYPERLINK("http://www.google.com/", "Google")
    
    // file 2 upload
    var blob2 = form.updloadFile2;    
    var file2 = folder.createFile(blob2);    
    file2.setDescription("Uploaded by " + form.memberName);
    var file2Url = file2.getUrl();
    //Logger.log(file2Url);
    sheet.getRange(sheet.getLastRow(), 27).setFormula('=hyperlink("' + file1Url + '","Image")');
    sheet.getRange(sheet.getLastRow(), 29).setFormula('=hyperlink("' + file2Url + '","Image")');
    
    // insert thumbnail image
    RFPsheet.setRowHeight(31, 150);// define a row height to determine the size of the image in the cell
    RFPsheet.setRowHeight(32, 150);
    RFPsheet.setColumnWidth(7, 150)
    RFPsheet.getRange(31, 7).setFormula('=image("http://drive.google.com/uc?export=view&id='+file1.getId()+'")');
    RFPsheet.getRange(32, 7).setFormula('=image("http://drive.google.com/uc?export=view&id='+file2.getId()+'")');
    
    // add 5 minute delay to allow images to propagate into spreadsheet
    // this approach didn't work, just had the form hanging for 5 minutes, then propagated images/emailed pdfs
    // delay(form.memberName,form.memberEmail,form.suppliersPOC);
    
    // add delay
    // this approach didn't work, just had the form hanging for 1 minutes, then propagated images/emailed pdfs  
    // Utilities.sleep(60 * 1000);
    
    // paste in the data
    RFPsheet.getRange(3,3).setValue(form.memberName);
    
    // get the RFP specifc template sheet
    // RFP template sheet must not be hidden in RFP web app master sheet
    var destination = SpreadsheetApp.openById('1VIXyIKRUgTCpj7B8_FGE27LV8_jKC0FYIbQXKUwKSr0');
    SpreadsheetApp.setActiveSpreadsheet(destination);
    RFPsheet.copyTo(destination);
    
    // delete old RFP form first
    var deletedSheet = SpreadsheetApp.setActiveSheet(destination.getSheets()[0]);
    //Logger.log(deletedSheet.getName());
    destination.deleteActiveSheet();
    
    // migrate data to RFP form
    dataIntoPDF(form.memberName,form.memberEmail,form.suppliersPOC);
    
    return "RFP from submitted successfully";
    
    } catch (error) {
    
    return error.toString();
  }
}


// delay function for images in pdf
// this function is NOT used
function delay(email1,email2,email3) {
  
  // add 5 minute delay before calling pdf/email function
  Utilities.sleep(300 * 1000);
  
  dataIntoPDF(email1,email2,email3);
}

// get supplier names
function getAdminInfo() {
  
  // get the Category sheet
  var categorySheet = SpreadsheetApp.openById('1kL5RSCuLtpgLDrMWDT5WMWFKVPceTHiNs7EW0ljq_Z4').getSheetByName('Categories');
  
  // get the Suppliers sheet
  var supplierSheet = SpreadsheetApp.openById('1kL5RSCuLtpgLDrMWDT5WMWFKVPceTHiNs7EW0ljq_Z4').getSheetByName('Suppliers');
  
  // create ranges of categories, supplier names and POC names
  // create category array
  // create supplier and POC name array
  // combine into one array to return
  
  var categoryArray = categorySheet.getRange(2,1,categorySheet.getLastRow()-1,1).getValues();
  
  var supplierArray = supplierSheet.getRange(2,1,supplierSheet.getLastRow()-1,2)  
    .getValues().reduce(function(p,c) {
      p.push(c);
      return p;
    },[]);
  
  var infoArray = {array1: categoryArray,
                   array2: supplierArray};
  
  //Logger.log(infoArray);
  
  return infoArray;
  /*
  
  
  return infoArray;
  */
}

// transfer data into RFP page and email to relevant parties
function dataIntoPDF(name,email1,email2) {
  
  Logger.log("Success");
  
  // get the RFP sheet
  var doc = SpreadsheetApp.openById('1VIXyIKRUgTCpj7B8_FGE27LV8_jKC0FYIbQXKUwKSr0');
  var RFPsheet = doc.getSheets()[0];
  
  /* Email RFP to relevant parties */
  
  // select addresses (make this auto-populate based on form entries)
  var adminEmail = "benlcollins@gmail.com";
  var memberEmail = email1;
  var supplierEmail = email2;
  
  // subject and time of email
  var subject = "RFP PDF Report " + (new Date()).toString();
  
  // body of email
  var body = "Email containing PDF generated from RFP report " + (new Date()).toString();
  
  // get the URL and remove word edit from end
  var url = doc.getUrl().replace(/edit$/,'');
  
  // create url extension
  var url_ext = 'export?exportFormat=pdf&format=pdf'   // export as pdf
                + '&size=letter'                       // paper size
                + '&portrait=false'                    // orientation, false for landscape
                + '&fitw=true&source=benlcollins'           // fit to width, false for actual size
                + '&sheetnames=false&printtitle=false' // hide optional headers and footers
                + '&pagenumbers=false&gridlines=false' // hide page numbers and gridlines
                + '&fzr=false'                         // do not repeat row headers (frozen rows) on each page
                + '&gid=';                             // the sheet's Id
  
  // OAuth token
  var token = ScriptApp.getOAuthToken();
  
  Logger.log(url);
  Logger.log(url_ext);
  Logger.log(token);
  
  // Convert worksheet to PDF
  var response = UrlFetchApp.fetch(url + url_ext + RFPsheet.getSheetId(), {
      headers: {
        // need to figure out what is happening with this oauth2 - works with token from my other script, but not one generated in this app
        'Authorization': 'Bearer ' + token
      }
    });
  
  var pdfResponse = response.getBlob().setName(RFPsheet.getName() + '.pdf');
  
  Logger.log(pdfResponse.getBytes());
  
  // send the email with the PDF attachment
  MailApp.sendEmail(adminEmail, subject, body, {attachments:[pdfResponse]});
  MailApp.sendEmail(memberEmail, subject, body, {attachments:[pdfResponse]});
  MailApp.sendEmail(supplierEmail, subject, body, {attachments:[pdfResponse]});
  
  /* final step - call the method to clear out the RFP form */
  clearRFP();
}



// clear out RFP form
function clearRFP() {
 // code in here to set all the RFP cells to "" 
  Logger.log("Calls clear out function successfully");
}