/* Create the needed custom menus on sheet
*/
function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Update Menu')
      .addItem('Calculate LCCF', 'updateLCCF')
      .addItem('See LC Salden','LCSalden')
      .addItem('See Fines (Strafen)','updateStrafen')
      .addToUi();
   SpreadsheetApp.getUi()
  .createMenu('Exact')
      .addItem('Import Values', 'ImportValues') 
      .addToUi();
}

/**
*   This method imports Exact values for the financial year 2016 (aka FINBalanceGLAAccountSearch) using the Exact CSV API. 
*   URL restriction: cvsdemiliter must be ";", order of cvscolumns should not change,exportlines should be 0 in order to 
*                    import all accounts, division is the division number of a mandat (@Exact URL) und partner key should not
*                    be changed, rcaction is restricted from Exact should always be 31 and sysexporting also 4.
*   @Warning : get rid of NULL bytes sent by server.
*/

function ImportValues() {
  
  //Exact URL for file
  //TODO: replace with correct URL.
  var theUrl ="EXACT-URL";
 
  //Fetch data from url  
  var response = UrlFetchApp.fetch(theUrl);
  var payload = response.getContentText();
  var bytes= response.getContent();
  var myArray = []
  //Get rid of NULL bytes
  for ( var i =1 ; i<bytes.length; i++){
    if (bytes[i] != 0){
      myArray.push(bytes[i]);
    }
  }
 
  var t = Utilities.newBlob(myArray).getDataAsString();      
    
  //Split all accounts and assign to arrayData
  var arrayData = []; 
  arrayData = t.split("\"");  

  //Build real array of data without "" and semicolons. Different characters encoding!
  var realArray= [];
  var semicolon = arrayData[2];
  var space = arrayData[10];
  for (var j=1 ; j< arrayData.length; ++j){
    if (arrayData[j] != semicolon  && arrayData[j] != space && arrayData[j] !=" "){
      realArray.push(arrayData[j]);
    }      
  }
  
  //Build 2 dimensional array and assign values
  var rowSplit = [];
  for (i=0 ; i< realArray.length;++i){
    rowSplit[i] = [];
  }  
  var counter = 0;
  var number = 0;
  for (var i = 0; i < realArray.length; ++i) {
    for (var j =0 ; j < 5; ++j){
      rowSplit[counter].push(realArray[number]);
      number++;
    }
    counter++;
  }
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();    
  var mysheet = spreadsheet.getSheetByName("Current values");
  var length = rowSplit.length;
  mysheet.getRange( "A1:E" +length).setValues(rowSplit); 
}
