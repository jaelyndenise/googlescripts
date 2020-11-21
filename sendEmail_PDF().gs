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
