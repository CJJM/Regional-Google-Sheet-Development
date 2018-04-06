/*
 * Underlying object for sheet utility
 * methods are included through 'prototype'
 */
function SheetUtils_() {}

/*
 * Object we will interact with!
 *
 * Example:
 *   SheetUtils.clearSheet('Sheet 1')
 */
var SheetUtils = new SheetUtils_()

SheetUtils_.prototype.getTime = function() {
  var date = new Date();
  var monthNames = [
    "Jan", "Feb", "Mar",
    "Apr", "May", "June",
    "July", "Aug", "Sept",
    "Oct", "Nov", "Dec"
  ];

  var day = date.getDate();
  var month = date.getMonth();
  var hour = date.getHours();
  var ampm = hour >= 12 ? 'PM' : 'AM';
  hour = hour % 12;
  hour = hour ? hour : 12; // the hour '0' should be '12'  
  
  var minute = ('0'+date.getMinutes()).slice(-2);

  var dateString = String(monthNames[month] + '-' + day + " " + hour + ":" + minute + " " + ampm);
  return dateString;
}


SheetUtils_.prototype.getCategory = function(categoryInt) {
  var category;
  switch (categoryInt) {
    case 1:
      category = "Revenue";
      break;
    case 2:
      category = "Bookings";
      break;      
    case 3:
      category = "PMO";
      break;
    case 4:
      category = "Operations";
      break;      
    case 5:
      category = "RMO";
      break;      
  }
  return category;
}


SheetUtils_.prototype.getPrecision = function(precisionInt) {
  var precision;
  switch (precisionInt) {
    case 1:
      precision = "Normal";
      break;
    case 2:
      precision = "High-level Summary";
      break;      
    case 3:
      precision = "Low-level Details";
      break;
  }
  return precision;
}
     
// Get column numbers and sheet name of Note sheet
SheetUtils_.prototype.getNoteSheetInfo = function() {
  var noteArray = [];
  var configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  noteArray[0] = configSheet.getRange(1, 2).getValue();
  for(var i = 2; i < 15; i++) {
    noteArray[(i-1)] = +configSheet.getRange(i, 2).getValue();
  }
  Logger.log("Note sheet array values are " + noteArray);
  return noteArray;
}
     
// Get column numbers and sheet name of Summary sheet
SheetUtils_.prototype.getMainSheetInfo = function() {
  var summaryArray = [];
  var configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  summaryArray[0] = configSheet.getRange(1, 4).getValue();
  for(var i = 2; i < 10; i++) {
    summaryArray[(i-1)] = +configSheet.getRange(i, 4).getValue();
  }
  Logger.log("Main sheet array values are " + summaryArray);
  return summaryArray;
}
     
SheetUtils_.prototype.getTaskSheetInfo = function() {
  var taskArray = [];
  var configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  taskArray[0] = configSheet.getRange(1, 6).getValue();
  for(var i = 2; i < 12; i++) {
    taskArray[(i-1)] = +configSheet.getRange(i, 6).getValue();
  }
  Logger.log("Task sheet array values are " + taskArray);
  return taskArray;
}  







     
     
     
     
     