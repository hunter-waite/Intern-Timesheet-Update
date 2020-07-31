function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Intern Menu')
      .addItem('Update Confirmed Intern Timesheets', 'menuItem1')
      .addItem('Clear Approvals', 'menuItem2')
      .addItem('Approve All', 'menuItem3')
      .addToUi();
}

function menuItem1() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Intern Timesheets");
  var templates = sheet.getRange("K3:K17");
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
    sheet.getRange(initRow + i, 12).setValue(newSheet.getRange("A1").getValue());
    if(templates[i] == "Hazard Mapping") {
      
      //var newRow = [["Invoiced", "Quarter", "Intern", "Date", "Detail", "Hours", "Rate", "Expenses"]];
      var newRow = [["",                                           // Invoiced
                     "",                                           // Quarter
                     "H. Waite",                                   // Intern
                     sheet.getRange(initRow + i, 4).getValue(),    // Date
                     sheet.getRange(initRow + i, 3).getValue(),    // Detail
                     sheet.getRange(initRow + i, 7).getValue(),    // Hours
                     sheet.getRange(initRow + i, 8).getValue(),    // Rate
                     sheet.getRange(initRow + i, 9).getValue()]];  // Expenses
      newSheet.insertRowAfter(10);
      newSheet.getRange(11,1, 1, 8).setValues(newRow);
    }
  }
}

function menuItem2() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Intern Timesheets");
  sheet.getRange("J3:J17").setValue(null);
}

function menuItem3() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Intern Timesheets");
  sheet.getRange("J3:J17").setValue('Yes');
}