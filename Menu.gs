/*
 * Used to create the menu that is used
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Intern Menu')
      .addItem('Import Data', 'menuItem1')
      .addItem('Push to Templates', 'menuItem2')
      .addSeparator()
      .addItem('Approve All', 'menuItem3')
      .addItem('Clear Approvals', 'menuItem4')
      .addToUi();
}

/* 
 * Function that loops through the interns and calls for an import of their data
 */
function menuItem1() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Intern Timesheets");
  // location of the interns tab
  var interns = sheet.getRange("Q3:S19").getValues();
  
  // Loops through the interns and checks to see if they need a fetch
  for (var i = 0; i < interns.length; i++) {
    var url = interns[i][1];
    var destRange = interns[i][2];
    if(url == "Timesheet URL" || interns[i][0] == "New Intern" || interns[i][0] == "")
      continue;
    
    importInternData(url, destRange);
  }
  // sets all to yes then clears in an effort to refresh 
}

/*
 * Loops through all the interns and adds their data to the correct templates
 */
function menuItem2() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Intern Timesheets");
  var interns = sheet.getRange("Q3:T19").getValues();
  
  for (var i = 0; i < interns.length; i++) {
    if(interns[i][0] == "" || interns[i][1] == "" || interns[i][2] == "" || interns[i][3] == "")
      continue;
    Logger.log(interns[0])
    updateTemplate(interns[i][0], interns[i][3]);
  }
  sortSheets();
}

/*
 * Sets all the times for interns to approved
 */
function menuItem3() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Intern Timesheets");
  sheet.getRange("J3:J17").setValue('Yes');
  sheet.getRange("J23:J37").setValue("Yes");
  sheet.getRange("J43:J57").setValue("Yes");
  sheet.getRange("J63:J77").setValue("Yes");
  sheet.getRange("J83:J97").setValue("Yes");
  sheet.getRange("J103:J117").setValue("Yes");
  sheet.getRange("J123:J137").setValue("Yes");
  sheet.getRange("J143:J157").setValue("Yes");
  sheet.getRange("J163:J177").setValue("Yes");
  sheet.getRange("J183:J197").setValue("Yes");
  sheet.getRange("J203:J217").setValue("Yes");
}

/*
 * Clears all the approval ratings for the interns
 */
function menuItem4() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Intern Timesheets");
  sheet.getRange("J3:J17").setValue(null);
  sheet.getRange("J23:J38").setValue(null);
  sheet.getRange("J43:J57").setValue(null);
  sheet.getRange("J63:J77").setValue(null);
  sheet.getRange("J83:J97").setValue(null);
  sheet.getRange("J103:J117").setValue(null);
  sheet.getRange("J123:J137").setValue(null);
  sheet.getRange("J143:J157").setValue(null);
  sheet.getRange("J163:J177").setValue(null);
  sheet.getRange("J183:J197").setValue(null);
  sheet.getRange("J203:J217").setValue(null);
}

/*
 * Using the url and the destination range pulls data from the URL spreadhseet
 * and puts it in the current spreadsheet
 */
function importInternData(url, destRangeString) {
  var destSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Intern Timesheets");
  var sourceSheet = SpreadsheetApp
    .openByUrl(url)
    .getSheetByName("Timesheet_Template");
  sourceRange = sourceSheet.getRange("A10:I25").getValues();
  destRange = destSheet.getRange(destRangeString);
  destRange.setValues(sourceRange);
}

/*
 * Grabs the dtat from the specified intern and template range
 * Loops through all the days and puts in correct template
 * Inserts at the top of correct template
 * Clears the approved rows as it is parsed through
 */
function updateTemplate(intern, templateRange) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Intern Timesheets");
  var templates = sheet.getRange(templateRange);
  var initRow = templates.getRow();
  var newSheet = null;
  templates = templates.getValues();
  for ( var i = 0 ; i < templates.length ; i++ ) {
    var currRow = "A" + (initRow + i) + ":" + "J" + (initRow + i);
    if(templates[i] == 'No Template' || templates[i] == "" || 
       templates[i] == "Not Approved Yet" || sheet.getRange(initRow + i, 9).getValue() == "") {
      continue;
    }
    newSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(templates[i]);
    if(newSheet == -1) {
      continue;
    }
    var newRow = [["",                                           // Invoiced
                   "",                                           // Quarter
                   intern,                                       // Intern
                   sheet.getRange(initRow + i, 4).getValue(),    // Date
                   sheet.getRange(initRow + i, 3).getValue(),    // Detail
                   sheet.getRange(initRow + i, 7).getValue(),    // Hours
                   sheet.getRange(initRow + i, 8).getValue(),    // Rate
                   sheet.getRange(initRow + i, 9).getValue()]];  // Expenses
    newSheet.insertRowBefore(11);
    newSheet.getRange(11,1, 1, 8).setValues(newRow);
    sheet.getRange(currRow).clearContent();
  }
}

function sortSheets() {
  var sapSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Intern Timesheets");
  var sapValues = sapSheet.getRange("O3:O19").getValues();
  
  var sortCol = 4; //D column (date)
  
  for ( var i = 0; i < sapValues.length ; i++ ) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sapValues[i]);
    if ( sheet == -1 || sheet ==  null  )
      continue;
    
    var range = sheet.getRange("A11:H");
  
    range.sort( { column : sortCol, ascending : false } );
  }
}

