function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Search')
      .addItem('Search benlcollins','searchBenlcollins')
      .addItem('Search Google','searchGoogle')
      .addToUi();
}


function searchGoogle() {
  
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
  
  //Browser.msgBox(keyword);
  
  // https://www.googleapis.com/customsearch/v1?parameters
  // must have 3 parameters:
  // api key
  // custom search engine id
  // search query
  // e.g. GET https://www.googleapis.com/customsearch/v1?key=INSERT_YOUR_API_KEY&cx=017576662512468239146:omuauf_lfve&q=lectures
  
  var parameters = {
    "muteHttpExceptions": true
  };
  
  var url = "https://www.googleapis.com/customsearch/v1?key=" + API_KEY + "&cx=" + GOOGLE_SEARCH_ENGINE_ID + "&q=" + keyword;
  
  try {
    var response = UrlFetchApp.fetch(url, parameters).getContentText();
    var json = JSON.parse(response);
    var items = json["items"];
    
    items.forEach(function(item) {
      Logger.log(item["htmlTitle"]);
    });
    
  }
  catch (e) {
    Logger.log(e);
  };
  
}


function searchBenlcollins() {
  
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
  
  //Browser.msgBox(keyword);
  
  // https://www.googleapis.com/customsearch/v1?parameters
  // must have 3 parameters:
  // api key
  // custom search engine id
  // search query
  // e.g. GET https://www.googleapis.com/customsearch/v1?key=INSERT_YOUR_API_KEY&cx=017576662512468239146:omuauf_lfve&q=lectures
  
  var parameters = {
    "muteHttpExceptions": true
  };
  
  var url = "https://www.googleapis.com/customsearch/v1?key=" + API_KEY + "&cx=" + BEN_SEARCH_ENGINE_ID + "&q=" + keyword;
  
  try {
    var response = UrlFetchApp.fetch(url, parameters).getContentText();
    var json = JSON.parse(response);
    var items = json["items"];
    
    items.forEach(function(item) {
      Logger.log(item["htmlTitle"]);
    });
    
    
    //Logger.log(json["items"]);
    
  }
  catch (e) {
    Logger.log(e);
  };
  
}