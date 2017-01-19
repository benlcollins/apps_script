// another namespace created in a different gs file
// different gs files irrelevant, all compile to one global space
var otherHelpers = (function() {
  
  var ns = {};
  
  ns.left = function(someString,startNum,numDigits) {
    return someString.substr(startNum,numDigits);
  }
  
  return ns;
  
})();