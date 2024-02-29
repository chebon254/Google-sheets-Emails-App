function sendBulkEmails() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();
    var uniqueEmails = {}; // Object to store unique emails
    var duplicates = []; // Array to store duplicate emails
  
    // Iterate through each row to remove duplicates
    data.forEach(function(row) {
      var email = row[0];
      if (uniqueEmails[email]) {
        // If email already exists, add it to duplicates array
        duplicates.push(row);
      } else {
        // Store the email as key in the object to ensure uniqueness
        uniqueEmails[email] = row;
      }
    });
  
    // Clear existing data in the sheet
    sheet.clear();
    
    // Write header row back to the sheet
    sheet.appendRow(["Email", "Full Name", "First Name", "Last Name", "Validity"]);
  
    // Iterate through unique emails
    Object.values(uniqueEmails).forEach(function(row) {
      var email = row[0];
      var firstName = row[2];
      var validity = validateEmail(email);
      
      // Write validity status to column E
      sheet.appendRow([email, row[1], firstName, row[3], validity ? "valid" : "invalid"]);
  
      // If the email is valid, send email
      if (validity === true) {
        var body = `
          <h1>Welcome to Our Product!</h1>
          <p>Dear ${firstName},</p>
          <p>We are excited to introduce you to our latest product. Click <a href="https://www.dummylink.com">here</a> to learn more and make a purchase!</p>
          <p>Thank you for choosing us!</p>
        `;
        
        MailApp.sendEmail({
          to: email,
          subject: "Welcome to Our Product!",
          htmlBody: body
        });
      }
    });
  
    // Create or get the duplicate sheet
    var duplicateSheet = getOrCreateDuplicateSheet();
    
    // Write duplicate emails to the duplicate sheet
    duplicates.forEach(function(row) {
      duplicateSheet.appendRow([row[0]]);
    });
  
    // Resize columns
    sheet.autoResizeColumns(1, sheet.getLastColumn());
    duplicateSheet.autoResizeColumns(1, duplicateSheet.getLastColumn());
    
    // Update formatting for the spreadsheet
    updateSpreadsheetFormatting();
  }
  
  // Function to validate email format
  function validateEmail(email) {
    var re = /\S+@\S+\.\S+/;
    return re.test(email);
  }
  
  // Function to get or create a sheet for duplicates
  function getOrCreateDuplicateSheet() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var duplicateSheet = spreadsheet.getSheetByName("Duplicates");
    if (!duplicateSheet) {
      duplicateSheet = spreadsheet.insertSheet("Duplicates");
      duplicateSheet.appendRow(["Duplicate Emails"]);
    }
    return duplicateSheet;
  }
  
  function updateSpreadsheetFormatting() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = spreadsheet.getSheets();
  
    // Update formatting for each sheet
    sheets.forEach(function(sheet) {
      updateSheetFormatting(sheet);
    });
  }
  
  function updateSheetFormatting(sheet) {
    // Set font to Montserrat and size to 10
    var font = "Montserrat";
    var size = 10;
    var range = sheet.getDataRange();
    range.setFontFamily(font).setFontSize(size);
  
    // Adjust column widths
    var lastColumn = sheet.getLastColumn();
    for (var col = 1; col <= lastColumn; col++) {
      var maxContentLength = 0;
      var values = sheet.getRange(1, col, sheet.getLastRow(), 1).getValues();
      values.forEach(function(row) {
        var content = row[0].toString().length;
        if (content > maxContentLength) {
          maxContentLength = content;
        }
      });
      sheet.setColumnWidth(col, maxContentLength * 7); // Adjust the multiplier as needed for appropriate width
    }
  }
  
  

function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Email Tools')
  .addItem('Send Message', 'sendBulkEmails')
  .addToUi();
}

