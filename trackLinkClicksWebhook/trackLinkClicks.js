/**
 * doPost function webhook to track link clicks
 */
function doPost(e) {

  // get spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Sheet1');
  
  // proceed if e exists
  if (typeof e !== 'undefined') {
    
    // parse webhook data to extract parameters
    const name = e.parameter.name;
    const section = e.parameter.section;
    const link = e.parameter.link;
    
    // create new timestamp
    const d = new Date();

    // append new row of data to Sheet
    sheet.appendRow([d, section, name, link]);
    
    // return undefined
    return;
  }
}
