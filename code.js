function processEmails() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Step 1: Delete duplicate emails
  removeDuplicates(sheet);
  
  // Step 2: Validate emails and add status
  validateEmails(sheet);
  
  // Step 3: Send message from Google Sheets
  sendMessage(sheet);
}

// Function to remove duplicate emails
function removeDuplicates(sheet) {
  let emailColumn = sheet.getRange("A:A").getValues();
  let uniqueEmails = [...new Set(emailColumn.flat())]; // Flatten the 2D array and get unique values
  
  sheet.getRange("A:A").clearContent(); // Clear the entire email column
  
  uniqueEmails.forEach(email => {
    if (email !== "") {
      sheet.appendRow([email]); // Append unique emails back to the sheet
    }
  });
}

// Function to validate emails and add status
function validateEmails(sheet) {
  let lastRow = sheet.getLastRow();
  let emails = sheet.getRange("A1:A" + lastRow).getValues();
  let statusColumn = sheet.getRange("B1:B" + lastRow); // Next column for status
  
  let statuses = [];
  
  emails.forEach((email, index) => {
    if (isValidEmail(email[0])) {
      statuses.push(["Valid"]);
    } else {
      statuses.push(["Invalid"]);
    }
  });
  
  statusColumn.setValues(statuses);
}

// Function to check if an email is valid
function isValidEmail(email) {
  // Simple regex for email validation
  let emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

// Function to send message from Google Sheets
function sendMessage(sheet) {
  let lastRow = sheet.getLastRow();
  let emailColumn = sheet.getRange("A1:A" + lastRow).getValues();
  let subject = "Your Subject Here";
  
  emailColumn.forEach(email => {
    if (isValidEmail(email[0])) {
      let body = `
        <h1>Welcome to Our Product!</h1>
        <p>Dear Client,</p>
        <p>We are excited to introduce you to our latest product. Click <a href="https://www.dummylink.com">here</a> to learn more and make a purchase!</p>
        <p>Thank you for choosing us!</p>
      `;
      MailApp.sendEmail(email[0], subject, "", { htmlBody: body });
    }
  });
}
