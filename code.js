function sendBulkEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var uniqueEmails = {}; // Object to store unique emails
  
  // Iterate through each row to remove duplicates
  data.forEach(function(row) {
    var email = row[0];
    // Store the email as key in the object to ensure uniqueness
    uniqueEmails[email] = row;
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
}

// Function to validate email format
function validateEmail(email) {
  var re = /\S+@\S+\.\S+/;
  return re.test(email);
}
