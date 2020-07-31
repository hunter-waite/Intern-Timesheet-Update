function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Intern Menu')
      .addItem('Import Data', 'menuItem1')
      .addItem('Update Confirmed Intern Timesheets', 'menuItem2')
      .addItem('Approve All', 'menuItem3')
      .addItem('Clear Approvals', 'menuItem4')
      .addToUi();
}

function menuItem1() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Intern Timesheets");
  var interns = sheet.getRange("Q3:S19").getValues();
  
  for (var i = 0; i < interns.length; i++) {
    var url = interns[i][1];
    var destRange = interns[i][2];
    if(url == "Timesheet URL" || interns[i][0] == "New Intern" || interns[i][0] == "")
      continue;
    
    importInternData(url, destRange);
  }
  menuItem3();
  menuItem4();
}

function menuItem2() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Intern Timesheets");
  var interns = sheet.getRange("Q3:T19").getValues();
  
  for (var i = 0; i < interns.length; i++) {
    if(interns[i][0] == "" || interns[i][1] == "" || interns[i][2] == "" || interns[i][3] == "")
      continue;
    Logger.log(interns[0])
    updateTemplate(interns[i][0], interns[i][3]);
  }
}

function menuItem3() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Intern Timesheets");
  sheet.getRange("J3:J17").setValue('Yes');
  sheet.getRange("J23:J37").setValue("Yes");
  sheet.getRange("J43:J57").setValue("Yes");
}

function menuItem4() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Intern Timesheets");
  sheet.getRange("J3:J17").setValue(null);
  sheet.getRange("J23:J38").setValue(null);
  sheet.getRange("J43:J57").setValue(null);
}

function importInternData(url, destRangeString) {
  var destSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Intern Timesheets");
  var sourceSheet = SpreadsheetApp
    .openByUrl(url)
    .getSheetByName("Timesheet_Template");
  sourceRange = sourceSheet.getRange("A10:I25").getValues();
  destRange = destSheet.getRange(destRangeString);
  destRange.setValues(sourceRange);
}

function updateTemplate(intern, templateRange) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Intern Timesheets");
  var templates = sheet.getRange(templateRange);
  var initRow = templates.getRow();
  var newSheet = null;
  templates = templates.getValues();
  for (var i = 0; i < templates.length; i++) {
    if(templates[i] == 'No Template' || templates[i] == "" || templates[i] == "Not Approved Yet") {
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
  }
}
