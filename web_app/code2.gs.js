// for publishing as web app
function doGet(e) {
  return HtmlService.createTemplateFromFile('index.html')
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}


// submit RFP form
function submitRFP(form) {
  //Logger.log("RFP submitted successfully!");
  
  // setup lock service
  var lock = LockService.getPublicLock();
  lock.waitLock(30000);  // wait 30 seconds before conceding defeat.
  
  try {
    
    /* recording content to spreadsheet section */
    // set where we write the data - you could write to multiple/alternate destinations
    var doc = SpreadsheetApp.openById('SHEET_ID');
    var sheet = doc.getSheetByName('Results');
    var RFPsheet = doc.getSheetByName('RFPasPDF');
    
    // clear out the RFP sheet
    clearRFP(RFPsheet);
    
    var timestamp = new Date();
    
    var innerData = [timestamp,
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
                 form.itemCost2,
                 form.itemName3,
                 form.itemDesc3,
                 form.itemQuantity3,
                 form.itemCost3,
                 form.itemName4,
                 form.itemDesc4,
                 form.itemQuantity4,
                 form.itemCost4,
                 form.itemName5,
                 form.itemDesc5,
                 form.itemQuantity5,
                 form.itemCost5,
                 form.itemName6,
                 form.itemDesc6,
                 form.itemQuantity6,
                 form.itemCost6,
                 form.itemName7,
                 form.itemDesc7,
                 form.itemQuantity7,
                 form.itemCost7,
                 form.itemName8,
                 form.itemDesc8,
                 form.itemQuantity8,
                 form.itemCost8,
                 form.itemName9,
                 form.itemDesc9,
                 form.itemQuantity9,
                 form.itemCost9,
                 form.itemName10,
                 form.itemDesc10,
                 form.itemQuantity10,
                 form.itemCost10
                ];

    // create array of main data and wrap in extra [] for apps script
    var mainData = [innerData.splice(0,17)];
    
    // create variable to store the starting row of new data
    var startRow = sheet.getLastRow()+1;
    
    // read the main data into the spreadsheet
    var outputRange = sheet.getRange(startRow, 1, 1, 17);
    outputRange.setValues(mainData.map(function(d){
      return d; 
    }));

    // read the itemized data into 10 rows of results spreadsheet
    var newArray = [innerData.splice(0,4),
                    innerData.splice(0,4),
                    innerData.splice(0,4),
                    innerData.splice(0,4),
                    innerData.splice(0,4),
                    innerData.splice(0,4),
                    innerData.splice(0,4),
                    innerData.splice(0,4),
                    innerData.splice(0,4),
                    innerData.splice(0,4)];
    
    //Logger.log(newArray);
    //Logger.log(newArray.length);
    
    var outputItemized = sheet.getRange(startRow,18,10,4);
    outputItemized.setValues(newArray);
    
    // formula to multiply the item quantity and item description and add into sheet, for each row
    for (i = 0; i < 10; i++) {
      sheet.getRange(startRow+i, 22).setFormula('=T'+(startRow+i)+'*U'+(startRow+i));
    };

    
    /* uploading file section */
    var dropbox = "Test Images XYZ";
    var folder, folders = DriveApp.getFoldersByName(dropbox);
    
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(dropbox);
    }
    
    // upload files and add hyperlinks
    var blobs = [form.updloadFile1,
                 form.updloadFile2,
                 form.updloadFile3,
                 form.updloadFile4,
                 form.updloadFile5,
                 form.updloadFile6,
                 form.updloadFile7,
                 form.updloadFile8,
                 form.updloadFile9,
                 form.updloadFile10];
    
    for (var i = 0; i < 10; i++) {
      //Logger.log(blobs[i].length);
      
      if (blobs[i].length > 0) {
        var file = folder.createFile(blobs[i]);
        file.setDescription("Uploaded by " + form.memberName);
        var fileUrl = file.getUrl();
        sheet.getRange(startRow+i, 23).setFormula('=hyperlink("' + fileUrl + '","Image")');
        RFPsheet.setRowHeight(31+i, 150);// define a row height to determine the size of the image in the cell
        RFPsheet.getRange(31+i, 7).setFormula('=image("http://drive.google.com/uc?export=view&id='+file.getId()+'")');
      }; 
    };
    
    // color the new row to help identify where each new entry starts
    var newItemRow = sheet.getRange(startRow,1,1,23);
    newItemRow.setBackground("#dddddd");
    
    // paste main data to the RFP sheet
    var mainRFPData = sheet.getRange(startRow,1,1,17).getValues();
    //Logger.log(mainRFPData);
    var mainRFPArray = mainRFPData[0].concat();
    //Logger.log(mainRFPData);
    mainRFPArray.splice(0,1);
    mainRFPArray.splice(4,0,'','');
    mainRFPArray.splice(10,0,'','');
    mainRFPArray.splice(16,0,'','');
    mainRFPArray.splice(21,0,'');
    //Logger.log(mainRFPArray);
    
    RFPsheet.getRange(3,3,23,1).setValues(mainRFPArray.map(function(d) {
      //Logger.log(d);
      return [d];
    }));

    // paste itemized data to the RFP sheet
    var itemRFPData = sheet.getRange(startRow,18,10,5).getValues();
    //Logger.log(itemRFPData);
    //var itemRFPArray = itemRFPData[0].concat();
    //Logger.log(itemRFPArray);
    RFPsheet.getRange(31,2,10,5).setValues(itemRFPData);
    
    // get the RFP specifc template sheet
    // RFP template sheet must not be hidden in RFP web app master sheet
    var destination = SpreadsheetApp.openById('SHEET_ID');
    SpreadsheetApp.setActiveSpreadsheet(destination);
    RFPsheet.copyTo(destination);
    
    // delete old RFP form first
    var deletedSheet = SpreadsheetApp.setActiveSheet(destination.getSheets()[0]);
    //Logger.log(deletedSheet.getName());
    destination.deleteActiveSheet();
    
    // migrate data to RFP form
    dataIntoPDF(form.memberName,form.memberEmail,form.suppliersPOC);
    
    return "RFP form submitted successfully";
    
    } catch (error) {
    
    return error.toString();
  }
}


// get supplier names
function getAdminInfo() {
  
  // get the Category sheet
  var categorySheet = SpreadsheetApp.openById('SHEET_ID').getSheetByName('Categories');
  
  // get the Suppliers sheet
  var supplierSheet = SpreadsheetApp.openById('SHEET_ID').getSheetByName('Suppliers');
  
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

}

// transfer data into RFP page and email to relevant parties
function dataIntoPDF(name,email1,email2) {
  
  Logger.log("Success");
  
  // get the RFP sheet
  var doc = SpreadsheetApp.openById('SHEET_ID');
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
  
  // Convert worksheet to PDF
  var response = UrlFetchApp.fetch(url + url_ext + RFPsheet.getSheetId(), {
      headers: {
        // need to figure out what is happening with this oauth2 - works with token from my other script, but not one generated in this app
        'Authorization': 'Bearer ' + token
      }
    });
  
  var pdfResponse = response.getBlob().setName(RFPsheet.getName() + '.pdf');
  
  // send the email with the PDF attachment
  MailApp.sendEmail(adminEmail, subject, body, {attachments:[pdfResponse]});
  MailApp.sendEmail(memberEmail, subject, body, {attachments:[pdfResponse]});
  MailApp.sendEmail(supplierEmail, subject, body, {attachments:[pdfResponse]});
  
  /* final step - call the method to clear out the RFP form */
  clearRFP();
}



// clear out RFP form
function clearRFP(sheet) {
  sheet.getRange(3,3,23,1).clearContent();
  sheet.getRange(31,2,10,5).clearContent();
  sheet.getRange(31,7,10,1).clearContent();
}