function sendBulkEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.autoResizeColumns(1, 10);
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
  sheet.appendRow(["Email", "Full Name", "First Name", "Last Name", "Validity", "Sent Status"]);

  // Iterate through unique emails
  Object.values(uniqueEmails).forEach(function(row) {
    var email = row[0];
    var firstName = row[2];
    var validationResult = verifyEmail(email);
    var validity = validationResult.status;

    // Write validity status to column E
    sheet.appendRow([email, row[1], firstName, row[3], validity]);

    // If the email is valid, send email
    if (validity === 'success') {
      var body = `
        <h1>Welcome to Our Product!</h1>
        <p>Dear ${firstName},</p>
        <p>We are excited to introduce you to our latest product. Click <a href="https://www.dummylink.com">here</a> to learn more and make a purchase!</p>
        <p>Thank you for choosing us!</p>
      `;

      // Print "Sent" in column three
      sheet.getRange(sheet.getLastRow(), 6).setValue("Sent");
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

  // Resize columns for both main sheet and duplicate sheet
  //doResizeColumns(sheet);
  //doResizeColumns(duplicateSheet);
  // Update formatting for the spreadsheet
  updateSpreadsheetFormatting();

  // Save the spreadsheet in the specified folder structure on Google Drive
  saveSpreadsheetToDrive();
}


function doResizeColumns(sheet){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  
  // Iterate through each sheet in the spreadsheet
  sheets.forEach(function(sheet) {
    // Auto resize columns for each sheet
    sheet.autoResizeColumns(1, 10);
  });
}



function verifyEmail(email) {
  var options = {
    method: 'GET',
    followRedirects: true,
    validateHttpsCertificates: false // This option bypasses SSL certificate validation
  };
  
  var url = "https://api.eva.pingutil.com/email?email=" + encodeURIComponent(email);
  
  try {
    var response = UrlFetchApp.fetch(url, options);
    var result = JSON.parse(response.getContentText());
    
    // Check if status is success and other conditions are met
    if (result.status === 'success' && result.data.deliverable === true) {
      return { status: 'success' };
    } else {
      return { status: 'failure' };
    }
  } catch (error) {
    console.error('Error:', error);
    return { status: 'failure', error: error };
  }
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
  
  // Function to update formatting for the sheet
  function updateSheetFormatting(sheet) {
    // Set font to Montserrat and size to 10
    var font = "Montserrat";
    var size = 10;
    var range = sheet.getDataRange();
    
    // Set font family and size
    range.setFontFamily(font).setFontSize(size);
    
    // Set font weight to bold
    range.setFontWeight("regular");

  }

  
  
  // Function to save the spreadsheet in the specified folder structure on Google Drive
  function saveSpreadsheetToDrive() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var folderName = "Email Sorted"; // Parent folder name
    var parentFolder = DriveApp.getFoldersByName(folderName);
    var parentFolderId;
    
    // Check if parent folder exists, if not create it
    if (!parentFolder.hasNext()) {
      var newFolder = DriveApp.createFolder(folderName);
      parentFolderId = newFolder.getId();
    } else {
      parentFolderId = parentFolder.next().getId();
    }
    
    // Get current timestamp for folder naming
    var now = new Date();
    var timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MM-yyyy_HH-mm");
    var dayOfWeek = Utilities.formatDate(now, Session.getScriptTimeZone(), "EEEE-d-MM-yyyy");
    var timeOfDay = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm:ss");
    
    // Create child folders if they don't exist
    var dayFolder = DriveApp.getFolderById(parentFolderId).getFoldersByName(dayOfWeek);
    var dayFolderId;
    if (!dayFolder.hasNext()) {
      var newDayFolder = DriveApp.getFolderById(parentFolderId).createFolder(dayOfWeek);
      dayFolderId = newDayFolder.getId();
    } else {
      dayFolderId = dayFolder.next().getId();
    }
    
    var timeFolder = DriveApp.getFolderById(dayFolderId).getFoldersByName(timeOfDay);
    var timeFolderId;
    if (!timeFolder.hasNext()) {
      var newTimeFolder = DriveApp.getFolderById(dayFolderId).createFolder(timeOfDay);
      timeFolderId = newTimeFolder.getId();
    } else {
      timeFolderId = timeFolder.next().getId();
    }
    
    // Save the spreadsheet in the child folder
    var ssFile = DriveApp.getFileById(spreadsheet.getId());
    var destinationFolder = DriveApp.getFolderById(timeFolderId);
    ssFile.makeCopy(ssFile.getName(), destinationFolder);
  }
  
  function onOpen() {
    SpreadsheetApp.getUi()
    .createMenu('Email Tools')
    .addItem('Send Message', 'sendBulkEmails')
    .addToUi();
  }
  