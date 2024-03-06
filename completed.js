// Function to send bulk emails
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

  // Write header row back to the sheet with additional columns for tracking
  sheet.appendRow(["Email", "Full Name", "First Name", "Last Name", "Validity", "Sent Status", "Email Opened", "Unsubscribe Hash"]);

  // Iterate through unique emails
  Object.values(uniqueEmails).forEach(function(row) {
    var email = row[0];
    var firstName = row[2];
    var unsubscribeHash = getMD5Hash(email);
    var validationResult = verifyEmail(email);
    var validity = validationResult.status;

    // Write validity status to column E
    sheet.appendRow([email, row[1], firstName, row[3], validity, '', '', unsubscribeHash]);

    // If the email is valid, send email
    if (validity === 'success') {
      var body = `
      <!DOCTYPE html>
      <html lang="en">
      <head>
          <meta charset="UTF-8">
          <meta http-equiv="X-UA-Compatible" content="IE=edge">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
      
          <!-- == Favicon == -->
          <link rel="shortcut icon" href="https://www.gucci.com/_ui/responsive/common/images/favicon.png">
      
          <title>GUCCI®</title>
      
          <style>
              @font-face {
                  font-family: '29LT Bukra';
                  src: url("29LT Bukra Medium.otf") format("opentype");
              }
              @font-face {
                  font-family: 'Gotham';
                  src: url("GothamMedium.ttf") format("truetype");
              }
              @font-face {
                  font-family: 'Gotham Book';
                  src: url("GothamBook.ttf") format("truetype");
              }
              @font-face {
                  font-family: 'MADE SAONARA 2';
                  src: url("MADE SAONARA 2 PERSONAL USE.otf") format("opentype");
              }
              @font-face {
                  font-family: 'Futura Bk BT';
                  src: url("FuturaBookBT.ttf") format("truetype");
              }
          </style>
      
          <!-- == ICONS == -->
          <link rel="stylesheet" target="_blank" href="https://use.fontawesome.com/releases/v5.13.0/css/all.css" integrity="sha384-Bfad6CLCknfcloXFOyFnlgtENryhrpZCe29RTifKEixXQZ38WheV+i/6YWSzkz3V" crossorigin="anonymous">
          <link rel="stylesheet" href="../../Sites/fontawesome/css/all.css">
      
      </head>
      <body style="min-height: 100vh !important; width: 100vw !important; display: flex !important; align-items: center !important; justify-content: center !important; background-color: #f5f5f5;">
          <main style="min-height: 600px; max-width: 600px !important; flex-grow: 1 !important; background-color: #ffffff !important; padding: 6px 0px !important;">
              <div class="top" style="height: 60px !important; width: 95% !important; background-color: #000000 !important; margin: auto !important; display: flex !important; align-items: center !important; justify-content: center !important;">
                  <p style="width: fit-content !important; cursor: pointer !important; font-size: 26px !important; font-family: '29LT Bukra'; font-weight: 500 !important; color: #ffffff !important; text-align: center !important; letter-spacing: 0.2em; text-transform: uppercase;margin: 0px !important;" onclick="window.open('https://www.gucci.com/uk/en_gb/st/capsule/gucci-bamboo-1947?gclsrc=aw.ds&gclid=Cj0KCQjwzqSWBhDPARIsAK38LY-Fmtn6T0Lpzb_XGT0gwIAmfFjsYel87dGN1zEHOEKqagFQoF4y5cwaAvmyEALw_wcB&gclsrc=aw.ds')">GUCCI</p>
              </div>
              <div class="back-image" style="height: 250px !important; width: 96% !important; margin: 20px auto 0px !important; background-image: url('./back.png') !important; background-size: cover !important; background-repeat: no-repeat !important; display: flex !important; align-items: center !important; justify-content: center !important;">
                  <div class="text" style="height: fit-content !important; width: fit-content !important; padding: 20px !important; background-color: #ffffff99 !important; text-align: center !important;">
                      <h3 style="margin: 0 !important; font-family: 'MADE SAONARA 2'; font-size: 28px !important; letter-spacing: 0.22em; ">BAMBOO HANDLE BAGS</h3>
                      <p style="text-align: center !important; font-size: 14px !important;  font-family: 'Gotham'; font-weight: 500 !important; line-height: 16px !important; margin: 10px 0px 12px !important;">The bag's bamboo handles are expertly curved using a<br>flame-a process originally conceived by Florentine<br> during postwar period.</p>
                      <button style="height: 38px !important; width: 210px !important; font-family: 'Gotham'; background-color: #BCB7B6 !important; font-size: 14px !important; color: #ffffff !important; border: 2px solid #ffffff !important; outline: none !important;" onclick="window.open('https://www.gucci.com/uk/en_gb/st/capsule/gucci-bamboo-1947?gclsrc=aw.ds&gclid=Cj0KCQjwzqSWBhDPARIsAK38LY-Fmtn6T0Lpzb_XGT0gwIAmfFjsYel87dGN1zEHOEKqagFQoF4y5cwaAvmyEALw_wcB&gclsrc=aw.ds')">COLLECT YOUR BAG</button>
                  </div>
              </div>
              <div class="image-anime" style="width: 96% !important; margin: auto !important; display: flex !important; justify-content: space-between !important; align-items: center !important; position: relative !important; padding-bottom: 10px !important;">
                  <img src="animation.gif" alt="" srcset="" width="100%">
                  <div class="anime-left" style="height: 492px !important; text-align: center; width: 260px !important; position: absolute !important; right: 16px !important; top: 12px !important;">
                      <h3 style="font-size: 24px !important; font-weight: 400 !important; font-family:'Times New Roman', Times, serif !important; letter-spacing: -0.025em; text-transform: uppercase; margin: 0 !important;">Small top handle<br>bag with bamboo</h3>
                      <div class="underline" style="height: 2px !important; width: 100% !important; background-color: #BDB5B9 !important; margin: 6px 0px 10px !important;"></div>
                      <span style="font-size: 13px !important; font-family: 'Gotham Book'; font-style: normal; font-weight: 325; letter-spacing: -0.015em;">Choose your leather’s colour:</span>
                      <div class="btns" style="display: flex !important; align-items: center !important; justify-content: center !important; width: fit-content !important; margin: 10px auto !important;">
                          <button style="height: 24px !important; width: 24px; margin: 0px 5px !important; background-color: #DE9C3A !important; border-radius: 15px !important; border: 0 !important; outline: none !important;"></button>
                          <button style="height: 24px !important; width: 24px; margin: 0px 5px !important; background-color: #000000 !important; border-radius: 15px !important; border: 0 !important; outline: none !important;"></button>
                          <button style="height: 24px !important; width: 24px; margin: 0px 5px !important; background-color: #5E372F !important; border-radius: 15px !important; border: 0 !important; outline: none !important;"></button>
                          <button style="height: 24px !important; width: 24px; margin: 0px 5px !important; background-color: #007B00 !important; border-radius: 15px !important; border: 0 !important; outline: none !important;"></button>
                          <button style="height: 24px !important; width: 24px; margin: 0px 5px !important; background-color: #ffffff !important; border-radius: 15px !important; border: 0 !important; outline: none !important;" onclick="window.open('https://www.gucci.com/uk/en_gb/st/capsule/gucci-bamboo-1947?gclsrc=aw.ds&gclid=Cj0KCQjwzqSWBhDPARIsAK38LY-Fmtn6T0Lpzb_XGT0gwIAmfFjsYel87dGN1zEHOEKqagFQoF4y5cwaAvmyEALw_wcB&gclsrc=aw.ds')"><i class="fas fa-plus" style="font-size: 20px !important; font-weight: 800 !important;"></i></button>
                      </div>
                      <span style="font-size: 13px !important; font-family: 'Gotham Book'; font-style: normal; font-weight: 325; letter-spacing: -0.015em;">Make this forever yours. Add your initials!</span> <br>
                      <button style="height: 38px !important; width: 210px !important; background-color: #ffffff !important; margin-top: 12px !important; border: 1px solid #BCB7B6 !important; font-size: 13px !important; font-family: 'Gotham Book'; font-style: normal; font-weight: 325; text-transform: uppercase !important; padding: 0px !important;" onclick="window.open('https://www.gucci.com/uk/en_gb/st/capsule/gucci-bamboo-1947?gclsrc=aw.ds&gclid=Cj0KCQjwzqSWBhDPARIsAK38LY-Fmtn6T0Lpzb_XGT0gwIAmfFjsYel87dGN1zEHOEKqagFQoF4y5cwaAvmyEALw_wcB&gclsrc=aw.ds')">Continue shopping</button>
                  </div>
              </div>
              <div class="bag-image" style="width: 95% !important; margin: auto !important;">
                  <img src="bag root.png" alt="" width="100% !important">
              </div>
              <div class="services">
                  <p style="font-size: 24px !important; color: #000000 !important; text-align: center !important;">CLIENT SERVICES</p>
                  <div class="services-btns" style="display: flex !important; align-items: center !important; justify-content: center !important;">
                      <button style="height: 40px !important; width: 160px !important; outline: none !important; padding-left: 40px !important; border: 0 !important; margin: 0px 6px; background: #343434 !important; font-size: 12px !important; font-family: 'Futura Bk BT'; position: relative; font-style: normal !important; color: #ffffff !important; text-align: left !important;" onclick="window.open('https://www.gucci.com/uk/en_gb/st/capsule/gifts-services')">Gucci Services <img src="gift.png" alt="" style="width: 30px !important; position: absolute !important; left: 8px; top: 6px !important;"></button>
                      <button style="height: 40px !important; width: 160px !important; outline: none !important; padding-left: 40px !important; border: 0 !important; margin: 0px 6px; background: #343434 !important; font-size: 12px !important; font-family: 'Futura Bk BT'; position: relative; font-style: normal !important; color: #ffffff !important; text-align: left !important;" onclick="window.open('https://www.gucci.com/uk/en_gb/st/capsule/book-an-appointment')">Book An Appoinment <img src="gift.png" alt="" style="width: 30px !important; position: absolute !important; left: 8px; top: 6px !important;"></button>
                  </div>
              </div>
              <div class="helping-hand" style="width: 100% !important;">
                  <p style="font-size: 24px !important; color: #000000 !important; text-align: center !important;">NEED A HELPING HAND?</p>
                  <div class="services-btns" style="display: flex !important; align-items: center !important; justify-content: center !important;">
                      <button style="height: 40px !important; width: 160px !important; outline: none !important; border: 0 !important; margin: 0px 6px; background: #343434 !important; font-size: 12px !important; font-family: 'Times New Roman' !important; font-style: normal !important; font-weight: 400 !important; color: #ffffff !important;" onclick="window.open('https://api.whatsapp.com/send?phone=442074951445')">Whatsapp Us</button>
                      <button style="height: 40px !important; width: 160px !important; outline: none !important; border: 0 !important; margin: 0px 6px; background: #343434 !important; font-size: 12px !important; font-family: 'Times New Roman' !important; font-style: normal !important; font-weight: 400 !important;  color: #ffffff !important;" onclick="window.open('https://www.gucci.com/uk/en_gb/st/contact-us')">Call us at +1.877.482.2430</button>
                      <button style="height: 40px !important; width: 160px !important; outline: none !important; border: 0 !important; margin: 0px 6px; background: #343434 !important; font-size: 12px !important; font-family: 'Times New Roman' !important; font-style: normal !important; font-weight: 400 !important;  color: #ffffff !important;" onclick="window.open('https://www.gucci.com/uk/en_gb/st/contact-us')">Email Us</button>
                  </div>
                  <div class="logo" style="width: 100% !important; text-align: center !important; padding-top: 40px !important;">
                      <img src="gucci logo.png" alt=""  width="200px">
                  </div>
              </div>
              <div class="socials" style="width: 100% !important; padding: 40px 0px !important;">
                  <div class="social-icons" style="width: fit-content !important; margin: auto !important;">
                      <a style="font-size: 24px !important; color: grey !important; margin: 0px 16px !important" href="http://www.facebook.com/gucci" target="_blank"><i class="fab fa-facebook"></i></a>
                      <a style="font-size: 24px !important; color: grey !important; margin: 0px 16px !important" href="http://twitter.com/gucci" target="_blank"><i class="fab fa-twitter"></i></a>
                      <a style="font-size: 24px !important; color: grey !important; margin: 0px 16px !important" href="http://instagram.com/gucci" target="_blank"><i class="fab fa-instagram"></i></a>
                      <a style="font-size: 24px !important; color: grey !important; margin: 0px 16px !important" href="http://www.youtube.com/gucciofficial" target="_blank"><i class="fab fa-youtube"></i></a>
                      <a style="font-size: 24px !important; color: grey !important; margin: 0px 16px !important" href="https://play.google.com/store/apps/details?id=com.gucci.gucciapp" target="_blank"><i class="fab fa-google-plus"></i></a>
                      <a style="font-size: 24px !important; color: grey !important; margin: 0px 16px !important" href="https://pinterest.com/gucci/" target="_blank"><i class="fab fa-pinterest"></i></a>
                  </div>
              </div>
              <div class="bottom" style="height: 143px !important; background-color: #000000 !important; padding: 20px 0px !important; width: 100% !important;">
                  <p style="font-size: 14px !important; font-weight: 500 !important; font-family: 'Gotham Book' !important; color: #ffffff !important; text-align: center !important; margin: 14px 0px !important; line-height: 16px !important;">© 2016 - 2021 Guccio Gucci S.p.A. - All rights reserved. <br> G Commerce Europe S.p.A. - IT VAT nr 05142860484. SIAE LICENCE # <br> 2294/l/1936 and 5647/l/1936</p>
                  <p style="font-size: 14px !important; font-weight: 500 !important; font-family: 'Gotham Book' !important; color: #ffffff !important; text-align: center !important;"><a style="color: #ffffff !important;"  href="https://script.google.com/macros/s/AKfycbxyhzk8JpzP1S-vXp6UVAOtQzN9qKqHLaKxiHr2cZ6mLsZ7EJcG/exec?email=${encodeURIComponent(email)}&unsubscribe_hash=${unsubscribeHash}">Unsubscribe</a> | <a style="color: #ffffff !important;"  href="https://www.gucci.com/uk/en_gb/st/privacy-landing" target="_blank">Privacy Policy</a></p>
              </div>
          </main>
          <script>
      
          </script>
      </body>
      </html>
      `;

      // Print "Sent" in column six
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

// Function to handle unsubscribe request
function unsubscribeUser(emailToUnsubscribe, unsubscribeHash) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var emailIndex = data[0].indexOf('Email');
  var unsubscribeHashIndex = data[0].indexOf('Unsubscribe Hash');
  var subscribedIndex = data[0].indexOf('Subscribed');

  for (var i = 1; i < data.length; i++) {
    if (data[i][emailIndex] === emailToUnsubscribe && data[i][unsubscribeHashIndex] === unsubscribeHash) {
      sheet.getRange(i + 1, subscribedIndex + 1).setValue('no');
      return true;
    }
  }
  return false;
}

// Function to handle email opening and update status
function trackEmailOpening(email) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var emailIndex = data[0].indexOf('Email');
  var statusIndex = data[0].indexOf('Email Opened');

  for (var i = 1; i < data.length; i++) {
    if (data[i][emailIndex] === email) {
      sheet.getRange(i + 1, statusIndex + 1).setValue('Opened');
      return true;
    }
  }
  return false;
}

// Function to generate MD5 hash
function getMD5Hash(value) {
  value = value + generateRandomString(8); // Add randomness to the email
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, value, Utilities.Charset.UTF_8);
  var hash = '';

  for (var i = 0; i < digest.length; i++) {
    var byte = digest[i];
    if (byte < 0) byte += 256;
    var bStr = byte.toString(16);
    if (bStr.length == 1) bStr = '0' + bStr;
    hash += bStr;
  }

  return hash;
}

// Function to generate random string
function generateRandomString(length) {
  var randomNumber = Math.pow(36, length + 1) - Math.random() * Math.pow(36, length);
  var string = Math.round(randomNumber).toString(36).slice(1);
  return string;
}
  
  function setFixedWidthForColumns() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var fixedWidthMultiplier = 2; // Multiplier for the fixed width
    
    // Get the original width of the first six columns
    var originalWidths = [];
    for (var i = 1; i <= 9; i++) {
      originalWidths.push(sheet.getColumnWidth(i));
    }
    
    // Set the fixed width for the first six columns
    for (var i = 1; i <= 9; i++) {
      var originalWidth = originalWidths[i - 1];
      var fixedWidth = originalWidth * fixedWidthMultiplier;
      sheet.setColumnWidth(i, fixedWidth);
    }
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
      .addItem('Resize Columns', 'setFixedWidthForColumns')
      .addItem('Open Email Tracking Sidebar', 'openSidebar')
      .addToUi();
  }
  
  // Function to open the sidebar
  function openSidebar() {
    var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Email Tracking Sidebar')
      .setWidth(300);
    SpreadsheetApp.getUi().showSidebar(html);
  }
    