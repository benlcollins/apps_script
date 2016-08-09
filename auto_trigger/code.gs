// -----------------------------------------------------------------------------
// Auto Trigger system
// Ben Collins 2016
// -----------------------------------------------------------------------------


// -----------------------------------------------------------------------------
// add menu
// -----------------------------------------------------------------------------
function onOpen() { 
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu("Auto Trigger")
    .addItem("Run","runAuto")
    .addToUi();
}


// -----------------------------------------------------------------------------
// main function to control workflow - runs once
// -----------------------------------------------------------------------------
function runAuto() {
  
  // resets the loop counter if it's not 0
  refreshUserProps();
  
  // clear out the sheet
  clearData();
  
  // create trigger to run program automatically
  createTrigger();
}


// -----------------------------------------------------------------------------
// function to add new number to sheet
// called by trigger once per each iteration of loop
// -----------------------------------------------------------------------------
function addNumber() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Data');
  
  // get loop counter
  var userProperties = PropertiesService.getUserProperties();
  var loopCounter = Number(userProperties.getProperty('loopCounter'));
  
  // some limit on the loop
  var limit = 3;
  
  // if loop counter < number of batches
  if (loopCounter < limit) {
    
    // see what the counter value is at the start of the loop
    Logger.log(loopCounter);
    
    // do stuff
    var num = Math.ceil(Math.random()*100);
    sheet.getRange(sheet.getLastRow()+1,1).setValue(num);
    
    // increment the properties service counter for the loop
    loopCounter = loopCounter + 1;
    userProperties.setProperty('loopCounter', loopCounter);
    
    // see what the counter value is at the end of the loop
    Logger.log(loopCounter);
  }
  else {
    // Log message to confirm loop is finished
    sheet.getRange(sheet.getLastRow()+1,1).setValue("Finished");
    Logger.log("Finished");
    
    // delete trigger because we've reached the end of the loop
    deleteTrigger();
    
  }
}


// -----------------------------------------------------------------------------
// create trigger to run addNumber every minute
// -----------------------------------------------------------------------------
function createTrigger() {
  
  // Trigger every 1 minute
  ScriptApp.newTrigger('addNumber')
      .timeBased()
      .everyMinutes(1)
      .create();
}


// -----------------------------------------------------------------------------
// function to delete triggers
// -----------------------------------------------------------------------------
function deleteTrigger() {
  
  // Loop over all triggers and delete them
  var allTriggers = ScriptApp.getProjectTriggers();
  
  for (var i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
  
}


// -----------------------------------------------------------------------------
// function to clear data
// -----------------------------------------------------------------------------
function clearData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Data');
  
  // clear out the matches and output sheets
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2,1,lastRow-1,1).clearContent();
  }
}


// -----------------------------------------------------------------------------
// reset loop counter to 0 in properties
// -----------------------------------------------------------------------------
function refreshUserProps() {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('loopCounter', 0);
}
