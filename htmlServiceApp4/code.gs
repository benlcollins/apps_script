// for publishing as web app
function doGet(e) {
  return HtmlService.createTemplateFromFile('index.html')
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}


// submit RFP form
function submitRFP(form) {
  Logger.log("Success!");
  
  // setup lock service
  var lock = LockService.getPublicLock();
  lock.waitLock(30000);  // wait 30 seconds before conceding defeat.
  
  try {
    
    /* recording content to spreadsheet section */
    // set where we write the data - you could write to multiple/alternate destinations
    var doc = SpreadsheetApp.openById('SHEET_FILE_ID');
    var sheet = doc.getSheetByName('Results');
    var RFPsheet = doc.getSheetByName('RFPasPDF');
    
    var data = [[form.memberName,
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
                 form.itemName,
                 form.itemDesc,
                 form.itemQuantity,
                 form.itemCost
                ]];
    
    var outputRange = sheet.getRange(sheet.getLastRow()+1, 1, 1, 20);
    outputRange.setValues(data.map(function(d){
      return d; 
    }));
    
    // formula to multiply the item quantity and item description and add into sheet
    sheet.getRange(sheet.getLastRow(), 21).setFormula('=S'+sheet.getLastRow()+'*T'+sheet.getLastRow());

    
    /* uploading file section */
    var dropbox = "Test Images XYZ";
    var folder, folders = DriveApp.getFoldersByName(dropbox);
    
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(dropbox);
    }
    
    var blob = form.updloadFile;    
    var file = folder.createFile(blob);    
    file.setDescription("Uploaded by " + form.memberName);
    var fileUrl = file.getUrl();
    Logger.log(fileUrl);
    sheet.getRange(sheet.getLastRow(), 22).setFormula('=hyperlink("' + fileUrl + '","Image")');  // e.g. =HYPERLINK("http://www.google.com/", "Google")
    
    // migrate data to RFP form
    // RFPsheet.getRange(3,3).setValue(form.memberName);
    dataIntoPDF(form.memberName);
    
    // insert thumbnail image
    RFPsheet.setRowHeight(31, 150);// define a row height to determine the size of the image in the cell
    RFPsheet.setColumnWidth(7, 150)
    RFPsheet.getRange(31, 7).setFormula('=image("http://drive.google.com/uc?export=view&id='+file.getId()+'")');
    
    return "RFP from submitted successfully";
    
    } catch (error) {
    
    return error.toString();
  }
}


// get supplier names
function getAdminInfo() {
  
  // get the Category sheet
  var categorySheet = SpreadsheetApp.openById('SHEET_FILE_ID').getSheetByName('Categories');
  
  // get the Suppliers sheet
  var supplierSheet = SpreadsheetApp.openById('SHEET_FILE_ID').getSheetByName('Suppliers');
  
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
function dataIntoPDF(data) {
  
  // get the RFP sheet
  var doc = SpreadsheetApp.openById('SHEET_FILE_ID');
  var RFPsheet = doc.getSheetByName('RFPasPDF');
  
  // paste in the data
  RFPsheet.getRange(3,3).setValue(data);
  
  /* Email RFP to relevant parties */
  
  // select addresses (make this auto-populate based on form entries)
  var email_1 = "benlcollins@gmail.com";
  
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
  
  // Convert worksheet to PDF
  var response = UrlFetchApp.fetch(url + url_ext + RFPsheet.getSheetId(), {
      headers: {
        // oauth
        'Authorization': 'Bearer ' + token
      }
    });
  
  var pdfResponse = response.getBlob().setName(RFPsheet.getName() + '.pdf');
  
  // send the email with the PDF attachment
  MailApp.sendEmail(email_1, subject, body, {attachments:[pdfResponse]});
  
  /* final step - call the method to clear out the RFP form */
  clearRFP();
}



// clear out RFP form
function clearRFP() {
 // code in here to set all the RFP cells to "" 
}
