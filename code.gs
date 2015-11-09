// Setup: 8 columns in the Spreadsheet with these Column Names:
// TASK | NAME	| WHO	| DUE DATE	| SENT	| active? (must be "yes" or "no")	| niceName	| AppScript HTML (taskDetails)

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var AssignTasksButton = [ {name: 'Assign Tasks', functionName: 'AssignTasks'}];
  var EmailTasksButton = [ {name: 'Email Tasks', functionName: 'EmailTasks'}];
  ss.addMenu('Assignment Emails', AssignTasksButton);
  ss.addMenu('Task Reminder Emails', EmailTasksButton);
}

// ASSIGN THOSE TASKS

function AssignTasks() {


  // Get active spreadsheet, find out the last row and last column
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var lastRow = ss.getLastRow();
  var lastColumn = ss.getLastColumn();

  // Select the range I want to pull data from. I want the entire sheet, until the last row and column.
  // The reason I use getRange(2,1) is because the very first row is my header. I don't need these range values for sending emails.
  var range = sheet.getRange(2, 1, lastRow, lastColumn);
  var values = range.getValues();
  
  
  for (var i = 0; i < values.length-1; i++) {

      var task = values[i][0];
      var who = values[i][1]; 
      var dueDate = values[i][2]; 
      var todayDate = Date();
      var sent = values[i][3];
      var activeFlag = values[i][4];
 
      //email config
      var niceName = values[i][5];
      var taskDetails = values[i][6];
      
      // Writes for loop results to log for debugging.
      var email = "For loop result: " + task + " " + who + " " + dueDate
      Logger.log(email);
    
    // Writes email
    if ( activeFlag == "yes" ) {
      MailApp.sendEmail({
        to: who,
        subject: "You have been assigned a task: " + task + " due on " + dueDate,
        htmlBody: niceName + ": " + task + " " + "Due: " + dueDate + "<br><br>" + taskDetails
      });
      Logger.log("MailApp result: " + task + " " + who + " " + dueDate);
    } else {
      Logger.log("MailApp result: No email sent - active flag is false")
    };
    
    
    // Writes to spreadsheet "Sent" column, which is hardcoded to D2
    if ( activeFlag == "yes" ) {
      var sentColRangeName = "D"+(2+i)
      var sentCol = sheet.getRange(sentColRangeName);
      sentCol.setValue("📤Sending...");
      Utilities.sleep(500);// pause in the loop for 100 milliseconds
      sentCol.setValue("😄Sent "+todayDate);
      Logger.log("Sent Col changed: " + sentColRangeName);
    } else {
      Logger.log("Sent Col changed: None ")
    }
  }
}


// EMAIL THOSE TASKS

function EmailTasks() {


  // Get active spreadsheet, find out the last row and last column
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var lastRow = ss.getLastRow();
  var lastColumn = ss.getLastColumn();

  // Select the range I want to pull data from. I want the entire sheet, until the last row and column.
  // The reason I use getRange(2,1) is because the very first row is my header. I don't need these range values for sending emails.
  var range = sheet.getRange(2, 1, lastRow, lastColumn);
  var values = range.getValues();
  
  
  for (var i = 0; i < values.length-1; i++) {

      var task = values[i][0];
      var who = values[i][1]; 
      var dueDate = values[i][2]; 
      var todayDate = Date();
      var sent = values[i][3];
      var activeFlag = values[i][4];
 
      //email config
      var niceName = values[i][5];
      var taskDetails = values[i][6];
      
      // Writes for loop results to log for debugging.
      var email = "For loop results: " + task + " " + who + " " + dueDate
      Logger.log(email);
    
    // Writes email
    if ( activeFlag == "yes" ) {
      MailApp.sendEmail({
        to: who,
        subject: task + ": You have a task due on " + dueDate,
        htmlBody: niceName + ": " + task + " " + "Due: " + dueDate + "<br><br>" + taskDetails
      });
      Logger.log("MailApp result: " + task + " " + who + " " + dueDate);
    } else {
      Logger.log("MailApp result: No email sent - active flag is false")
    };
    
    
    // Writes to spreadsheet "Sent" column, which is hardcoded to D2
    if ( activeFlag == "yes" ) {
      var sentColRangeName = "D"+(2+i)
      var sentCol = sheet.getRange(sentColRangeName);
      sentCol.setValue("📤Sending...");
      Utilities.sleep(500);// pause in the loop for 100 milliseconds
      sentCol.setValue("😄Sent "+todayDate);
      Logger.log("Sent Col changed: " + sentColRangeName);
    } else {
      Logger.log("Sent Col changed: None ")
    }
  }
}
