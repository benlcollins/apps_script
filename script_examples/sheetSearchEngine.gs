// https://cse.google.com/cse/all
// https://developers.google.com/custom-search/json-api/v1/using_rest
// https://support.google.com/customsearch//answer/2630963?hl=en&ctx=topic

// custom menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Search')
      .addItem('Search benlcollins','searchBenlcollins')
      .addItem('Search Google','searchGoogle')
      .addToUi();
}


// Google search function
function searchGoogle() {
  var results = searchFunction(GOOGLE_SEARCH_ENGINE_ID);
  
  var output = [];
  
  results.forEach(function(result) {
    output.push([result["htmlTitle"],result["link"],result["snippet"]]);
  });
  
  // Logger.log(output);
  outputResults(output);
}


// benlcollins site search
function searchBenlcollins() {
  searchFunction(BEN_SEARCH_ENGINE_ID);
}


// output results into sheet
function outputResults(output) {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // clear any old content
  sheet.getRange(4,1,sheet.getLastRow(),3).clearContent();
  
  // insert new information
  sheet.getRange(4,1,output.length,3).setValues(output);
  
}
  
  

// general search function
function searchFunction(search_engine_id) {
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var keyword = sheet.getRange(1,2).getValue();
  
  // encode URI components if any
  // https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/encodeURIComponent
  keyword = encodeURIComponent(keyword);
  
  // replace spaces with +
  // the /.../gi means peform a global, case-insensitive replacement, i.e. all spaces
  keyword = keyword.replace(/%20/gi, "+");
  
  // Remove any ?
  keyword = keyword.replace(/%3F/gi, "");
  
  // Google custom search
  // https://developers.google.com/custom-search/json-api/v1/using_rest
  // e.g. GET https://www.googleapis.com/customsearch/v1?key=INSERT_YOUR_API_KEY&cx=017576662512468239146:omuauf_lfve&q=lectures
  
  var parameters = {
    "muteHttpExceptions": true
  };
  
  var url = "https://www.googleapis.com/customsearch/v1?key=" + API_KEY + "&cx=" + search_engine_id + "&q=" + keyword;
  
  try {
    var response = UrlFetchApp.fetch(url, parameters).getContentText();
    var json = JSON.parse(response);
    var items = json["items"];
    Logger.log(items[2]);
    
    /*
    items.forEach(function(item) {
      Logger.log(item["htmlTitle"]);
    });
    */
    
    return items;
    
  }
  catch (e) {
    Logger.log(e);
  };
  
}


