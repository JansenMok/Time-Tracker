function startActivity() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var now = new Date();

  // ask user for the activity name
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("activity name:");

  var activity = response.getResponseText();

  // if left blank reuse last activity from column B
  if (activity === "") {
    if (lastRow > 1) { // check that there is a previous row
      activity = sheet.getRange(lastRow, 2).getValue();
    } else {
      activity = "unnamed activity"; // fallback if no previous entry
    }
  }

  // write new entry in the next empty row
  var newRow = lastRow + 1;
  sheet.getRange(newRow, 1).setValue(Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd")); // Date
  sheet.getRange(newRow, 2).setValue(activity); // Activity
  sheet.getRange(newRow, 3).setValue(now); // Start Time
}

function stopActivity() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var now = new Date();

  sheet.getRange(lastRow, 4).setValue(now); // End Time
}
