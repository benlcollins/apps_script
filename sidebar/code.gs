// apps script code to show sidebar of suppliers

// create menu on open
function onOpen(e) {
  
  SpreadsheetApp.getUi()
    .createMenu('Suppliers')
    .addItem('Choose Supplier','showSuppliers')
    .addToUi();
}

// execute the HTML service
function showSuppliers() {
  
  var ui = HtmlService.createTemplateFromFile('index.html')
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('Suppliers select demo');
  
  SpreadsheetApp.getUi().showSidebar(ui);
}

// get supplier names
function getSupplierNames() {
  // get the data in the active sheet
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // create a 2 dim range of supplier names
  return sheet.getRange(2,1,sheet.getLastRow()-1,2)
  .getValues().reduce(function(p,c) {
    p.push(c);
    return p;
  },[]);
}