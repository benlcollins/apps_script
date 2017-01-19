// function to call and log namespaces
function namespaces() {
 
  Logger.log(helpers.Ucase('are you sure'));  
  
  Logger.log(otherHelpers.left('testing this out',3,8));
  
}

// this is the namespace we create
var helpers = (function () {
  
  var ns = {};
  
  ns.Ucase = function(theString) {
    return theString.toUpperCase();
  }
  
  return ns;
})();
