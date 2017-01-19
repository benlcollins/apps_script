// ------------------------------------------------------------------------------------------
// Save social media data
// set to auto save with a trigger when dashboard is finalized
// ------------------------------------------------------------------------------------------
function saveSocialData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("social_media");
  
  // get the url, follower count and date from the first three cells
  var timestamp = sheet.getRange(14,2).getValue();
  var facebook_count = sheet.getRange(12,2).getValue();
  var twitter_count = sheet.getRange(12,3).getValue();
  var linkedin_count = sheet.getRange(12,4).getValue();
  var g_plus_count = sheet.getRange(12,5).getValue();
  var pinterest_count = sheet.getRange(12,6).getValue();
  var instagram_count = sheet.getRange(12,7).getValue();
  
  // paste them into the bottom row of your spreadsheet
  sheet.appendRow([
    timestamp,
    facebook_count,
    twitter_count,
    linkedin_count,
    g_plus_count,
    pinterest_count,
    instagram_count]);
};


// ------------------------------------------------------------------------------------------
// Save alexa data
// set to auto save with a trigger when dashboard is finalized
// ------------------------------------------------------------------------------------------
function saveAlexaData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("alexa_rankings");
  
  // get the url, follower count and date from the first three cells
  var timestamp = sheet.getRange(13,2).getValue();
  var global_count = sheet.getRange(11,2).getValue();
  var us_count = sheet.getRange(11,3).getValue();
  
  // paste them into the bottom row of your spreadsheet
  sheet.appendRow([
    timestamp,
    global_count,
    us_count]);
};




// ------------------------------------------------------------------------------------------
// FB likes code
// need to add error handling to this
// ------------------------------------------------------------------------------------------
function fb_likes(url,username) {
  
  var searchTerm = url + username;
  var apiURL = "http://api.facebook.com/restserver.php?method=links.getStats&urls=" + encodeURIComponent(searchTerm);
  
  try {
    var response = UrlFetchApp.fetch(apiURL).getContentText();
  }
  catch(e) {
    return "Error fetching data from Facebook.";
  }
  
  var xml = Xml.parse(response, false);
  var link_stat = xml.getElement().getElement();
  
  return parseInt(link_stat.getElement("like_count").getText());
};



// ------------------------------------------------------------------------------------------
// Twitter followers code
// ------------------------------------------------------------------------------------------
function twitter_followers(url,username) {
  
  var searchTerm = url + username;
  
  try {
    var html = UrlFetchApp.fetch(searchTerm).getContentText().toString();
  }
  catch(e) {
    return "Error fetching data from Twitter.";
  }
  
  var shift = 53 +  username.length; // length of the search string that we need to "shift" by
  var start = html.search('<a href="/' + username + '/followers">') + shift;  
  var html2 = html.substr(start);
  var end = html2.search('</div>');
  var answer = html2.substr(1,end-1);
  
  return answer;
}


// ------------------------------------------------------------------------------------------
// Linkedin followers code
// ------------------------------------------------------------------------------------------
// Still can't get linkedin working

function linkedin_followers(url,username) {

  //var test = "https://www.linkedin.com/in/benlcollins";
  var searchTerm = url + username;
  
  try {
    var html = UrlFetchApp.fetch(searchTerm).getContentText().toString();
  }
  catch(e) {
    return "Error fetching data from Linkedin.";
  }
  
  var start = html.search('member-connections');
  
  var html2 = html.substr(start);
  Logger.log(html2);
  //member-connections"><strong>266</strong> connections</div>
  
  var answer = html2.substring(html2.indexOf("<strong>")+8,html2.indexOf("</strong>"));
  Logger.log(answer);
  
  return answer;

}


// ------------------------------------------------------------------------------------------
// YouTube followers code
// ------------------------------------------------------------------------------------------
function youtube_followers(url,username) {

  var searchTerm = url + username;
  
  try {
    var html = UrlFetchApp.fetch(searchTerm).getContentText().toString();
  }
  catch(e) {
    return "Error fetching data from YouTube.";
  }
  
  var start = html.search('subscribers') + 12;
  
  var html2 = html.substr(start);
  
  var end = html2.search('</span>');
  
  var answer = html2.substr(1,end-1);
  Logger.log(answer);
  return answer;

}



// ------------------------------------------------------------------------------------------
// G+ followers code
// ------------------------------------------------------------------------------------------
function g_followers(url,username) {
  
  //var test = "https://plus.google.com/+benlcollins";
  
  var searchTerm = url + username;
  
  try {
    var html = UrlFetchApp.fetch(searchTerm).getContentText().toString();
  }
  catch(e) {
    return "Error fetching data from G+";
  }
  
  //Logger.log(html);
  
  var start = html.search('followers')-25;
  
  var html2 = html.substr(start,50);
  
  var answer = html2.substring(html2.indexOf(">")+1,html2.indexOf("<"));
  //Logger.log(answer);
  return answer;
}

// ------------------------------------------------------------------------------------------
// Pinterest followers code
// ------------------------------------------------------------------------------------------
function pinterest_followers(url,username) {

  var searchTerm = url + username;
  //var test = "https://www.pinterest.com/bencollins    09u09jiihvbb ";
  
  try {
    var html = UrlFetchApp.fetch(searchTerm).getContentText().toString();
  }
  catch(e) {
    return "Error fetching data from Pinterest";
    //throw "Error trying to fetch Pinterest data: " + e.message;
  }
  
  var start = html.search('FollowerCount');
  
  var html2 = html.substr(start);
  //Logger.log(html2);
  
  var answer = html2.substring(html2.indexOf("<span")+1,html2.indexOf("</span"));
  var answer2 = answer.substr(answer.indexOf(">")+1);
  
  Logger.log(answer2);
  
  return answer2;

}




// ------------------------------------------------------------------------------------------
// Instagram followers code
// ------------------------------------------------------------------------------------------
function instagram_followers(url,username){
  
  var searchTerm = url + username;
  
  try {
    var html = UrlFetchApp.fetch(searchTerm).getContentText().toString();
  }
  catch(e) {
    return "Error fetching data from Instagram.";
  }
  
  var start = html.search('followed_by')+23;
  var html2 = html.substr(start);
  var end = html2.search('}');
  
  var answer = html2.substr(1,end-1);
  return answer;
}
