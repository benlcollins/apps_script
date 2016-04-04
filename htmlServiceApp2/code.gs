// code to create menu
function onOpen(e) {
  
  SpreadsheetApp.getUi()
    .createMenu('Suppliers')
    .addItem('Choose Supplier','showSuppliers')
    .addItem('Use google.script.run','showSuppliersRun')
    .addItem('Use dialog','showDialog')
    .addToUi();
}


// for publishing as web app
function doGet(e) {
  return HtmlService.createTemplateFromFile('indexRun.html')
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}


// HTML service to create sidebar
function showSuppliers() {
  
  var ui = HtmlService.createTemplateFromFile('index.html')
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('Suppliers select demo');
  
  SpreadsheetApp.getUi().showSidebar(ui);
}

// HTML service to create sidebar without using HTML templating
function showSuppliersRun() {
  
  var ui = HtmlService.createTemplateFromFile('indexRun.html')
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('Suppliers select demo');
  
  SpreadsheetApp.getUi().showSidebar(ui);
}


// HTML service to create dialog popup box
function showDialog() {
  
  var html = HtmlService.createTemplateFromFile('indexRun.html')
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(400)
    .setHeight(300);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Suppliers Select Demo');
}
  

// get supplier names
function getSupplierNames() {
  // get the data in the active sheet
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // if there is not sheet, fall back to identifying the specific spreadsheet
  if (!sheet) {
    sheet = SpreadsheetApp.openById('SHEET_ID')
      .getSheetByName('Sheet1');
  }
  
  // create a 2 dim range of supplier names
  return sheet.getRange(2,1,sheet.getLastRow()-1,2)
  .getValues().reduce(function(p,c) {
    p.push(c);
    return p;
  },[]);
}
