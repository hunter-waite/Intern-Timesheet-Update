var sapLen = 16;

// Checks to see if manual validity check has been filled in yet
// fills out the cell next to it
function checkValidity(input) {
  
  // if the user has not approved
  if(input == "No") {  
    return("Not Approved Yet");
  }
  // if the user has not approved it yet, deafult to approved
  else {
    // get the sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Intern Timesheets");
    
    // gets row and column for the correct SAP order
    var currCell = sheet.getActiveCell();
    var firstColumn = 1;
    var currentRow = currCell.getRow();
    
    // gets sap order using row and column
    var sapOrder = sheet.getRange(currentRow, firstColumn).getValue();
    sapOrderName = getSheetName(sheet, sapOrder);
    
    return(sapOrderName);
  }
}

// Gets the sheet name associated with the SAP Order from the list
function getSheetName(sheet, input) {
  var location = getSapOrdersLocation(sheet);
  
  // gets the sap orders and associate name from the list of sap orders
  var sapOrders = sheet.getRange(location[0], location[1], sapLen, 2).getValues();
  var name = "No Template";
  
  // loops through to find name
  for (i = 0; i < sapLen; i++) {
    if (sapOrders[i][0] == input) {
      name = sapOrders[i][1];
      break;
    }
  }
  return(name);
}

// finds the location of the SAP Orders table in the spreadsheets
// upper left hand corner
function getSapOrdersLocation(sheet) {
  var sapName = 'SAP Orders';
  // gets all rows and columns
  var cols = sheet.getDataRange().getValues();
  // searches first row
  var sapColNum = cols[0].indexOf(sapName);
  var sapRowNum = 1;
  // loops throuygh rows until its found or past 100 rows
  while(sapColNum == -1 && sapRowNum < 100) {
    sapColNum = cols[sapRowNum].indexOf(sapName);
    sapRowNum++;
  }
  // returns row and column in form used by getRange()
  return([sapRowNum, sapColNum + 1]);
}