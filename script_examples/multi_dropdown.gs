// ----------------------------------------------------------------------
// code.gs file:
// ----------------------------------------------------------------------

// for publishing as web app
function doGet(e) {
  return HtmlService.createTemplateFromFile('index.html')
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}


// submit RFP form
function submitRFP(form) {
  
}


// get supplier names
function getCategorySupplier() {
  
  // get relevant sheet with data
  var sheet = SpreadsheetApp.openById('1gmW6vlxXBbsBynNdncXAFbtBcgcpO7CVwtL71a7SGCU').getSheetByName('supplier_categories');
  
  // get all category and supplier data from that sheet
  var values = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  
  // create an array of headings
  var headings = values.shift();
  var headerOb = makeHeaderOb(headings);
  var dataOb = makeDataOb (headerOb,values);
  Logger.log(dataOb[2].category);
  
  // make the headings array into an object
  // output:
  // {supplier a=2.0, supplier b=3.0, none=1.0, supplier c=4.0, category=0.0}
  function makeHeaderOb (head) {
    var idx = 0;
    return head.reduce(function(p,c) {
      p[c.toString().toLowerCase()] = idx++;
      return p;
    },{});
  }
  
  // make the values array into an object with the heading row
  function makeDataOb (headOb,values) {
      return values.map(function(row) {
          return Object.keys(headOb).reduce(function(p,c) {
            p[c] = row[headOb[c]];
            return p;
          },{});
      });
  }
  
  
  
  /*
  var categorySupplierObject = categorySupplierArray.reduce(function(o,v,i) {
    o[i] = v;
    return o;
  }, {});
  
  Logger.log(categorySupplierObject);
  
  */
  return dataOb;
}


// ----------------------------------------------------------------------
// render.js file
// ----------------------------------------------------------------------

<!-- render namespace is used to interact with the DOM - principally to build the options list based on
data retrieved from the server -->

<script>

// rendering happens here

var Render = (function(render) {

  /**
   * display the selection made on change
   * @param { object } e returned by change event
   */
  
  /**
   * build sector
   */
  render.build = function() {
    
    // get the category and supplier data
    var data = App.globals.data;
    
    
    // set the category and supplier selector options
    var categorySelector = document.getElementById("categories");
    var supplierSelector = document.getElementById("suppliers");
    
    /*
     * format of the data object at this stage:
     * [{supplier a=, supplier b=, none=X, supplier c=, category=Select category...}, 
     * {supplier a=X, supplier b=X, none=, supplier c=, category=Apps}, 
     * {supplier a=, supplier b=X, none=, supplier c=, category=AV Lighting & Production}, 
     * {supplier a=, supplier b=, none=, supplier c=X, category=Banners & Signs}, 
     * {supplier a=, supplier b=, none=X, supplier c=, category=Builder - Event Elements}]
    */
    
    // add the categories to the category select drop down
    data.forEach( function(d) {
      //console.log(d);
      categorySelector.appendChild(new Option (d.category, d.category));
    });
    
    // add an event listener and then add suppliers to the supplier select drop down
    categorySelector.addEventListener("change", function(e) {
      //want to clear out the current options
      removeOptions(supplierSelector);
    
      // add filtered list of suppliers
      // need to grab relevant row for selected category
      // need to map from X values to Supplier A, Supplier B etc.
      // i.e. where there is an X in a column, replace that X with the Supplier from that column
      // create new array with no blanks left in the array
      // append each element of the array as a child to the supplier select
    
      var chosenCategory = e.target.value;
      
      var chosenCategoryData = data.filter(function( obj ) {
        return obj.category == chosenCategory;
      })[0];
      
      var supplierList = Object.keys(chosenCategoryData).filter(function(key) {return chosenCategoryData[key] === "X"});
      //console.log(supplierList);
      
      supplierList.forEach(function(f) {
        supplierSelector.appendChild(new Option (f,f));
      });
    });
    
    
    // function to remove all options from a select element
    function removeOptions(selectbox) {
      var i;
      for(i=selectbox.options.length-1;i>=0;i--)
      {
        selectbox.remove(i);
      }
    }
  };
  
  return render;
  
})(Render || {});

</script>

