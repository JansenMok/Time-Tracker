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
  sheet.getRange(newRow, 1).setValue(Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd")); // date
  sheet.getRange(newRow, 2).setValue(activity); // activity
  sheet.getRange(newRow, 3).setValue(now); // start time

  // scroll down
  // sheet.setActiveRange(sheet.getRange(newRow, 1));
  sheet.getRange(newRow, 2).activate(); // activity name cell
  // SpreadsheetApp.flush(); // only if laggy
}

function stopActivity() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var now = new Date();

  sheet.getRange(lastRow, 4).setValue(now); // end time
  enforceTimeFormat(); // reformat just in case
}

// source table time formatter
function enforceTimeFormat() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // columns C and D are start and end time
  sheet.getRange("C:D").setNumberFormat("h:mm AM/PM");
}

function onOpen() {
  enforceTimeFormat();

  // jump to the latest activity on page load
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();              // or ss.getSheetByName("Sheet1")
  const last = sh.getLastRow();

  // if there’s at least one data row (header in row 1)
  if (last >= 2) {
    sh.activate();                              // ensure this sheet is focused
    sh.getRange(last, 2).activate();           // column B = Activity
    // SpreadsheetApp.flush(); // uncomment if you notice visual lag
  }
}

// this is too heavy
// const SOURCE_SHEET = 'tracker';
// function onEdit(e) {
//   if (!e) return;
//   const sh = e.range.getSheet();
//   if (sh.getName() !== SOURCE_SHEET) return;          // ignore other tabs

//   // If the edit intersects columns C:D (Start/End), re-apply the time format
//   const colStart = e.range.getColumn();
//   const colEnd   = colStart + e.range.getNumColumns() - 1;
//   const intersectsCD = !(colEnd < 3 || colStart > 4);
//   if (intersectsCD) {
//     sh.getRange(e.range.getRow(), Math.max(3, colStart),
//                 e.range.getNumRows(), Math.min(4, colEnd) - Math.max(3, colStart) + 1)
//       .setNumberFormat("h:mm AM/PM");
//   }
// }

// pivot table reformatter
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var colCount = sheet.getMaxColumns();

  // Example: force font & size for the last column if newly added
  var range = sheet.getRange(1, colCount, sheet.getMaxRows());
  range.setFontFamily("Victor Mono").setFontSize(14);
}

// pivot table bulk reformatter (run once)
// function onEdit(e) {
//   var sheet = e.source.getActiveSheet();
//   var range = sheet.getDataRange();
//   range.setFontFamily("Victor Mono").setFontSize(14);
// }
