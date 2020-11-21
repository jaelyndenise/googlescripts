# googlescripts
A compilation of google scripts (.gs)(Javascript(.js)) that can be used in a google sheets document. 

= = = = =
List of scripts and their functions:

1. sendEmail();
-this function sends an email to the address found in cell G18 that includes the invoice number written in cell B2. This script may be attached to a button or a trigger of your choosing.

function sendEmail() {
  var sheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  
  var invoiceNumberRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(1, 2); 
  var invoiceNumber = invoiceNumberRange.getValue();
  var ui = SpreadsheetApp.getUi(); 

  // Fetch the email address
  var emailRange = SpreadsheetApp.getActiveSheet().getRange(18, 7);
  var emailAddress = emailRange.getValue();

  var message = 'Thank you for submitting your work order form. Someone should be in touch with you soon. Your work order invoice number is: ' + invoiceNumber + "."; 
  var subject = 'VIVIDA Dermatology: Work Order Confirmation';
     
  MailApp.sendEmail(emailAddress, subject, message);
}


- - -
2. sendEmail_PDF();
-sends an email to the address found in cell G16 that includes the invoice number in cell B2 as well as a pdf attachment of the entire spreadsheet (not just the current sheet).

function sendEmail_PDF() {
  var sheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();

  var invoiceNumberRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(1, 2); 
  var invoiceNumber = invoiceNumberRange.getValue();
  
  
  var fileId = SpreadsheetApp.getActiveSpreadsheet().getId();
  

  var emailRange = SpreadsheetApp.getActiveSheet().getRange(16, 7);
  var emailAddress = emailRange.getValue();

  var message = 'Here is the .pdf file for Maintenance Invoice #: ' + invoiceNumber + "."; 
  var subject = 'Maintenance Invoice #: ' + invoiceNumber;
     
  MailApp.sendEmail(emailAddress, subject, message, {
        attachments: [DriveApp.getFileById(fileId)]
    });
};


= = =
Functions #3 - #8 require the following variables:

var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();

var offBoardSheet = ss.getSheetByName("OffBoarded");
var activeUserSheet = ss.getSheetByName("Assigned Key Cards");

- - -
3. replaceUsers();
-Trigger: Form Response 
     (please note, this function will not work properly unless this trigger is set. The function is built to isolate a single duplicate, not multiple duplicates - although, that wouldn't be terribly difficult to add in later).

-When a form response is submitted, this function finds any pre-existing card numbers (or entries) that match the newest entry, it copies the pre-existing entry to an "off-board" sheet, time stamps it for future reference, and removes it from the active sheet.

-This function allows users to keep track of active card numbers, retired numbers, and card transfers.
-This function requires functions #4 - #8 to be defined in the same script file.

function replaceUsers() {
  highlightDuplicates1();
  Logger.log("highlightDuplicates1(): Complete");
  
  copyHiLite();
  Logger.log("copyHiLite(): Complete");
             
  setOffBoardDate(); 
  Logger.log("setOffBoardDate(): Complete");
  
  removeHighlight();
  Logger.log("removeHighlight(): Complete");
  
  alignRight();
  Logger.log("alignRight(): Complete");
};

- - -
4. highlightDuplicates();
-Exception: This function IS a macro!
-This function finds the duplicates in column G and highlights them.

function highlightDuplicates1() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, 7, sheet.getMaxRows(), 1).activate();
  var conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getCurrentCell().offset(0, 0, 107, 1)])
  .whenCellNotEmpty()
  .setBackground('#FFE599')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getCurrentCell().offset(0, 0, 107, 1)])
  .whenFormulaSatisfied('=cou')
  .setBackground('#FFE599')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getCurrentCell().offset(0, 0, 107, 1)])
  .whenFormulaSatisfied('=countif')
  .setBackground('#FFE599')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getCurrentCell().offset(0, 0, 107, 1)])
  .whenFormulaSatisfied('=countif(G:G,G1)')
  .setBackground('#FFE599')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getCurrentCell().offset(0, 0, 107, 1)])
  .whenFormulaSatisfied('=countif(G:G,G1)>1')
  .setBackground('#FFE599')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getCurrentCell().offset(0, 0, 107, 1)])
  .whenFormulaSatisfied('=countif(G:G,G1)>1')
  .setBackground('#FFE599')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
};


- - -
5. copyHiLite();
-This function copies the row with the pre-existing card number to the "off-board" sheet, and then deletes that row from the "active" sheet.

function copyHiLite() {
  //Note: Only works if highlightDuplicates1() function has completed  
  for (c = 2; c < activeUserSheet.getLastRow(); c++) {    
    var color = activeUserSheet.getRange(c, 7).getBackground();
    if (color === "#ffe599") {
      activeUserSheet.getRange(c, 1, 1, 7).copyTo(offBoardSheet.getRange(offBoardSheet.getLastRow() + 1, 1, 1, 7)); 
      activeUserSheet.deleteRow(c);
      Logger.log("You're freaking awesome!");
      break;
    } else {
      Logger.log("replaceUser: Failed");
    };
  };
};


- - -
6. setOffBoardDate();
-Timestamps the row with the pre-existing card numbers that was copied on to the "off-board" sheet.

function setOffBoardDate() {
  var date = new Date();
  date.setHours(0, 0, 0, 0);
  offBoardSheet.getRange(offBoardSheet.getLastRow(), 8).setValue(date);
};


- - -
7. removeHighlight();
-This function removes the highlight from both the "active" and the "off-board" sheet by clearing the formatting.

function removeHighlight() {
  var current = sheet.getRange(2, 7, sheet.getMaxRows(), 1);
  var off = offBoardSheet.getRange(2, 7, offBoardSheet.getMaxRows(), 1);
  
  current.clear({formatOnly: true});
  off.clear({formatOnly: true})
};


- - -
8. alighRight();
-This function right-justifies all of the card numbers. The replaceUsers(); function will run without it but the spreadsheet just won't look as purdy).
-Exception: This IS a macro.

function alignRight() { 
  activeUserSheet.getRange(1, 7, sheet.getMaxRows(), 1).activate();
  ss.getActiveRangeList().setHorizontalAlignment('right');
  offBoardSheet.getRange(1, 7, sheet.getMaxRows(), 1).activate();
  ss.getActiveRangeList().setHorizontalAlignment('right');
};










