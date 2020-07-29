function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
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