function getProp() {
  var userProperties = PropertiesService.getUserProperties();
  return userProperties.getProperty('API_KEY');
}

var apiKey = getProp();