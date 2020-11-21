var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();

//var user = sheet.getRange(lastRow, 1
var offBoardSheet = ss.getSheetByName("OffBoarded");
var activeUserSheet = ss.getSheetByName("Assigned Key Cards");

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

function setOffBoardDate() {
  var date = new Date();
  date.setHours(0, 0, 0, 0);
  offBoardSheet.getRange(offBoardSheet.getLastRow(), 8).setValue(date);
};

function removeHighlight() {
  var current = sheet.getRange(2, 7, sheet.getMaxRows(), 1);
  var off = offBoardSheet.getRange(2, 7, offBoardSheet.getMaxRows(), 1);
  
  current.clear({formatOnly: true});
  off.clear({formatOnly: true})
};

function alignRight() { 
  activeUserSheet.getRange(1, 7, sheet.getMaxRows(), 1).activate();
  ss.getActiveRangeList().setHorizontalAlignment('right');
  offBoardSheet.getRange(1, 7, sheet.getMaxRows(), 1).activate();
  ss.getActiveRangeList().setHorizontalAlignment('right');
};

