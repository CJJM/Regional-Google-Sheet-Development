var NoteWindow;
var userProperties = PropertiesService.getUserProperties();
var documentProperties = PropertiesService.getDocumentProperties();

function takeProjectNote() {
  NoteWindow = HtmlService.createTemplateFromFile('NoteWindow').evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle("Notation Service")
      .setWidth(300);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(NoteWindow);
}

function showProjectNote() {
  DisplayNote = HtmlService.createTemplateFromFile('DisplayNote').evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle("View Note")
      .setWidth(300);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(DisplayNote);
}

function showNoteResponse(note) {
  var htmlTemplate = HtmlService.createTemplateFromFile('DisplayNote2');
  htmlTemplate.dataFromServerTemplate = note;
  var htmlOutput = htmlTemplate.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle("View Note")
      .setWidth(300);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(htmlOutput);
}

function showTaskMenu() {
  TaskWindow = HtmlService.createTemplateFromFile('TaskWriter').evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle("Assign Task")
      .setWidth(300);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(TaskWindow);
}

function saveUserProperties(rowValue, categoryValue, precisionValue, noteValue) {
  userProperties.setProperty("row", rowValue);
  userProperties.setProperty("category", categoryValue);
  userProperties.setProperty("precision", precisionValue);
  userProperties.setProperty("note", noteValue);
  Logger.log("Saving: " + rowValue +  " " + categoryValue + " " + precisionValue + " " + noteValue);
  writeWindowNote();
}

function getWindowProperties() {
  var props = {
    row:userProperties.getProperty("row"),
    category:userProperties.getProperty("category"), 
    precision:userProperties.getProperty("precision"), 
    note:userProperties.getProperty("note")
  };
  return props;
}


// experiment with caching values in Config sheet to speed up note taking job, test to see if difference is made by switching out 
function writeWindowNote() {
  Logger.log("Start");
  var noteSheetArray = SheetUtils.getNoteSheetInfo();
  var mainSheetArray = SheetUtils.getMainSheetInfo();  
  var noteSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(noteSheetArray[0]);
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(mainSheetArray[0]);
  var userString = String(Session.getActiveUser().getEmail());
  var noteLastRow = noteSheet.getLastRow();
  var nextRow = noteSheet.getLastRow() + 1;
  var date = SheetUtils.getTime();
  Logger.log("Made it here");
  var windowVals = getWindowProperties();

  var rowNumber = +windowVals.row;
  var category = SheetUtils.getCategory(+windowVals.category);  
  var precision = SheetUtils.getPrecision(+windowVals.precision);
  
  var note = "(" + date + " " + category + ") " + windowVals.note;
    
  // line item values from Current Quarter tab
  var revStatus = mainSheet.getRange(rowNumber, (mainSheetArray[1])).getValue();
  var opportunity = mainSheet.getRange(rowNumber, (mainSheetArray[2])).getValue();
  var project = mainSheet.getRange(rowNumber, (mainSheetArray[3])).getValue();
  var bookingStatus = mainSheet.getRange(rowNumber, (mainSheetArray[4])).getValue();
  var budgetStatus = mainSheet.getRange(rowNumber, (mainSheetArray[5])).getValue();
  var revTotal = mainSheet.getRange(rowNumber, (mainSheetArray[6])).getValue();
  var territorySalesManager = mainSheet.getRange(rowNumber, (mainSheetArray[7])).getValue();
  var projectNumber = mainSheet.getRange(rowNumber, (mainSheetArray[8])).getValue();
  
  noteSheet.getRange(nextRow, noteSheetArray[1]).setValue(date);
  noteSheet.getRange(nextRow, noteSheetArray[2]).setValue(opportunity);
  noteSheet.getRange(nextRow, noteSheetArray[3]).setValue(project);
  noteSheet.getRange(nextRow, noteSheetArray[4]).setValue(revStatus);
  noteSheet.getRange(nextRow, noteSheetArray[5]).setValue(bookingStatus);
  noteSheet.getRange(nextRow, noteSheetArray[6]).setValue(budgetStatus);
  noteSheet.getRange(nextRow, noteSheetArray[7]).setValue(precision);
  noteSheet.getRange(nextRow, noteSheetArray[8]).setValue(userString);
  noteSheet.getRange(nextRow, noteSheetArray[9]).setValue(category);
  noteSheet.getRange(nextRow, noteSheetArray[10]).setValue(note);
  noteSheet.getRange(nextRow, noteSheetArray[11]).setValue(revTotal);
  noteSheet.getRange(nextRow, noteSheetArray[12]).setValue(territorySalesManager);
  noteSheet.getRange(nextRow, noteSheetArray[13]).setValue(projectNumber);
  
  SpreadsheetApp.getUi().alert('Write function complete')
}

function retrieveNote(rowNumber, category, precision) {
//  var userProperties = PropertiesService.getUserProperties();
  var noteSheetArray = SheetUtils.getNoteSheetInfo();
  var mainSheetArray = SheetUtils.getMainSheetInfo();  
  var noteSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(noteSheetArray[0]);
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(mainSheetArray[0]);
  
  var opportunity = +mainSheet.getRange(rowNumber, mainSheetArray[2]).getValue();
  var project = mainSheet.getRange(rowNumber, mainSheetArray[3]).getValue();
  
  // TODO: edit how values are obtained here
  var precisionString = SheetUtils.getPrecision(+precision);
  var lastNoteRow = noteSheet.getLastRow();
  var lastMainRow = mainSheet.getLastRow();
  var noteRetrieved = false;
  var note; 
  
  Logger.log("Retrieve note got " + rowNumber + " " + category + " " + precision);
  // no longer need 
  // separate this into other function named checkLineItems
  for(var i = 1; i <= lastNoteRow; i++) {
    if(noteSheet.getRange(i, noteSheetArray[2]).getValue() == opportunity) {
      if(noteSheet.getRange(i, noteSheetArray[3]).getValue() == project) {
        if(noteSheet.getRange(i, noteSheetArray[7]).getValue() == precisionString) {
          if(noteRetrieved == true) {
            note = note + "   " + noteSheet.getRange(i, noteSheetArray[10]).getValue();
            Logger.log("Made it into note retrieval check");
            continue;
          }
          note = noteSheet.getRange(i, noteSheetArray[10]).getValue();
          Logger.log(noteRetrieved + " with a note of " + note);
          var noteRetrieved = true;
//          break;
        }
      }
    }
  }
//  if(noteRetrieved === true) {
//    for(var i = 1; i <= lastMainRow; i++) {
//      if(mainSheet.getRange(i, 10).getValue() == opportunity) {
//        if(mainSheet.getRange(i, 11).getValue() == project) {
//          mainSheet.getRange(i, 11).setNote(note);
//          break;
//        }
//      }
//    }
//  }
  showNoteResponse(note);
  return note;
}

function writeTaskTable(taskString, rowNumber, recipientEmail) {
  var taskSheetArray = SheetUtils.getTaskSheetInfo();  
  var mainSheetArray = SheetUtils.getMainSheetInfo();  
  var taskSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(taskSheetArray[0]);
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(mainSheetArray[0]);
  var lastMainRow = mainSheet.getLastRow();
  var lastTaskRow = taskSheet.getLastRow();
  var nextRow = lastTaskRow + 1;

  var date = SheetUtils.getTime();
  var userString = String(Session.getActiveUser().getEmail());
  // make this generic, add to config sheet
  var sheetHyperlink = "https://docs.google.com/spreadsheets/d/1es2wMyxW_yRp2OQoD5L-gZmkcbRzDdzqjoOJJByvBZo/edit?usp=sharing"
  var revStatus = mainSheet.getRange(rowNumber, (mainSheetArray[1])).getValue();
  var opportunity = mainSheet.getRange(rowNumber, (mainSheetArray[2])).getValue();
  var budget = mainSheet.getRange(rowNumber, (mainSheetArray[3])).getValue();
  var territorySalesManager = mainSheet.getRange(rowNumber, (mainSheetArray[7])).getValue();
  var projectNumber = mainSheet.getRange(rowNumber, (mainSheetArray[8])).getValue();

  taskSheet.getRange(nextRow, taskSheetArray[1]).setValue(date);
  taskSheet.getRange(nextRow, taskSheetArray[2]).setValue(opportunity);
  taskSheet.getRange(nextRow, taskSheetArray[3]).setValue(projectNumber);
  taskSheet.getRange(nextRow, taskSheetArray[4]).setValue(budget);
  taskSheet.getRange(nextRow, taskSheetArray[5]).setValue(revStatus);
  taskSheet.getRange(nextRow, taskSheetArray[6]).setValue(territorySalesManager);
  taskSheet.getRange(nextRow, taskSheetArray[7]).setValue(taskString);
  taskSheet.getRange(nextRow, taskSheetArray[8]).setValue(recipientEmail);
  taskSheet.getRange(nextRow, taskSheetArray[9]).setValue(userString);
  
  MailApp.sendEmail(recipientEmail, userString, "New Task To Complete", "Spreadsheet: " + sheetHyperlink  + "\nSheet: " + mainSheetArray[0] + "\nRow: " + rowNumber + "\n\nTask: " + taskString);
}

// create function with onEdit trigger to respond to assigner that task was completed



// Verify updated functionality works
function exportAll() {
  var noteSheetArray = SheetUtils.getNoteSheetInfo();
  var mainSheetArray = SheetUtils.getMainSheetInfo();  
  var noteSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(noteSheetArray[0]);
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(mainSheetArray[0]);
  var lastNoteRow = noteSheet.getLastRow();
  var lastMainRow = mainSheet.getLastRow();
  Logger.log("Last Note is at " + lastNoteRow + " and last summary row is at " + lastMainRow)
  
  for(var i = 3; i <= lastNoteRow; i++) {    
    var opp = noteSheet.getRange(i, noteSheetArray[2]).getValue();
    var project = noteSheet.getRange(i, noteSheetArray[3]).getValue();
    var revStatus = noteSheet.getRange(i, noteSheetArray[4]).getValue();
    var note = noteSheet.getRange(i, noteSheetArray[10]).getValue();
    Logger.log("Looking for the opp " + opp + " and project " + project);
    for(var j = 1; j <= mainSheet.getLastRow(); j++) {
      if(mainSheet.getRange(j, mainSheetArray[2]).getValue() == opp) {
        if(mainSheet.getRange(j, mainSheetArray[3]).getValue() == project) {
          Logger.log("Budget Line Item matched");
          mainSheet.getRange(j, mainSheetArray[3]).setNote(note);
        }
      }
    }
  }
}

function clearNotes() {
  var mainSheetArray = SheetUtils.getMainSheetInfo();  
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(mainSheetArray[0]);
  mainSheet.getRange(1, mainSheetArray[3], mainSheet.getLastRow()).clearNote();
}

function emailNotes() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Executive Log");
  var lastColumn = sheet.getLastColumn();
  var cell;
  var message = new Array;
  for(var i = 2; i <= lastColumn; i++) {
    cell = "" + sheet.getRange(6, i).getValue();
    message.push(cell);
  }
  MailApp.sendEmail("cmaloney@redhat.com", "Test Email", message);
}


