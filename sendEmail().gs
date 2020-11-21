function sendEmail() {
  var sheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  //var sheet = ss.getSheets();  
  //var sheetName = sheet.getNames().toString();
  
  var invoiceNumberRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(1, 2); 
  var invoiceNumber = invoiceNumberRange.getValue();
  var ui = SpreadsheetApp.getUi(); 

  // Fetch the email address
  var emailRange = SpreadsheetApp.getActiveSheet().getRange(18, 7);
  var emailAddress = emailRange.getValue();

  var message = 'Thank you for submitting your work order form. Someone should be in touch with you soon. Your work order invoice number is: ' + invoiceNumber + "."; 
  var subject = 'VIVIDA Dermatology: Work Order Confirmation';
     
  MailApp.sendEmail(emailAddress, subject, message);
};
