/**
 * ===================================================================
 * CONFIGURATION
 * ===================================================================
 * A. Set your preferences here.
 * B. Set up your sheet names and column *numbers* (A=1, B=2, etc.).
 * C. Create the email template in `emailTemplate.html`.
 */
const CONFIG = {
  /**
   * A. GENERAL SETTINGS
   */
  // The email address to send reminders from.
  // Session.getActiveUser().getEmail() uses the email of the person running the script.
  SENDER_EMAIL: Session.getActiveUser().getEmail(),
  
  // The display name for the sender.
  SENDER_NAME: "Your Team Name", // <-- CHANGE THIS
  
  // Your local timezone.
  // Full list: https://en.wikipedia.org/wiki/List_of_tz_database_time_zones
  TIMEZONE: "Asia/Jakarta", // <-- CHANGE THIS

  // The subject line for the email.
  // You can use ${recipientName} as a placeholder.
  EMAIL_SUBJECT_TEMPLATE: "Action Required: Missing Document for ${recipientName}", // <-- CHANGE THIS

  /**
   * B. SHEET & COLUMN CONFIGURATION (Use column *numbers*)
   */
  SHEET_VALIDATION: {
    NAME: "Validation",
    SENDER_COL: 2,           // B: Sender/Requestor Email
    RECIPIENT_NAME_COL: 3,   // C: Recipient Name
    RECIPIENT_EMAIL_COL: 6,  // F: Recipient Email
    ACTION_COL: 15,          // O: Action Dropdown
    STATUS_COL: 16           // P: Email Status
  },
  
  SHEET_REMINDER: {
    NAME: "Reminder Tool",
    EMAIL_COL: 1,            // A: Email Address
    NAME_COL: 2,             // B: Recipient Name
    TIMESTAMP_COL: 3,        // C: Last Send History
    STATUS_COL: 4,           // D: Send Status
    START_ROW: 4             // The first row containing data
  },
  
  /**
   * C. DROPDOWN & FORMATTING
   */
  DROPDOWN_OPTIONS: ["Send Email", "Email Sent"],
  DROPDOWN_ACTION_TEXT: "Send Email",
  DROPDOWN_SENT_TEXT: "Email Sent",
  DROPDOWN_ACTION_COLOR: "#FFD580", // Light Orange
  DROPDOWN_SENT_COLOR: "#006400"   // Dark Green
};

// ===================================================================
// END OF CONFIGURATION - DO NOT EDIT BELOW THIS LINE
// ===================================================================


/**
 * Creates the "Reminder Tool" menu when the spreadsheet is opened.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Reminder Tool")
    .addItem("Send Reminder Emails in Bulk", "sendReminderEmailsInBulk")
    .addItem("Setup Dropdown Tools", "setupDropdownTools")
    .addToUi();
}

/**
 * Sets up the dropdown and conditional formatting in the "Validation" sheet.
 */
function setupDropdownTools() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_VALIDATION.NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Error: Sheet '" + CONFIG.SHEET_VALIDATION.NAME + "' not found.");
    return;
  }

  // Set headers
  sheet.getRange(1, CONFIG.SHEET_VALIDATION.ACTION_COL).setValue("Action");
  sheet.getRange(1, CONFIG.SHEET_VALIDATION.STATUS_COL).setValue("Email Status");

  // Create dropdown rule
  var actionRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.DROPDOWN_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
    
  var maxRows = sheet.getMaxRows();
  sheet.getRange(2, CONFIG.SHEET_VALIDATION.ACTION_COL, maxRows - 1, 1).setDataValidation(actionRule);

  // Apply conditional formatting
  var oRange = sheet.getRange(2, CONFIG.SHEET_VALIDATION.ACTION_COL, maxRows - 1, 1);
  var ruleOrange = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(CONFIG.DROPDOWN_ACTION_TEXT)
    .setBackground(CONFIG.DROPDOWN_ACTION_COLOR)
    .setRanges([oRange])
    .build();
  var ruleDarkGreen = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(CONFIG.DROPDOWN_SENT_TEXT)
    .setBackground(CONFIG.DROPDOWN_SENT_COLOR)
    .setRanges([oRange])
    .build();

  var rules = sheet.getConditionalFormatRules();
  rules.push(ruleOrange, ruleDarkGreen); // Add new rules
  sheet.setConditionalFormatRules(rules);

  SpreadsheetApp.getUi().alert("Dropdown tools setup completed in '" + CONFIG.SHEET_VALIDATION.NAME + "' sheet.");
}

/**
 * Installable trigger function to handle edits in "Validation" sheet.
 */
function onEditTrigger(e) {
  var range = e.range;
  var sheet = range.getSheet();

  if (sheet.getName() === CONFIG.SHEET_VALIDATION.NAME) {
    var column = range.getColumn();
    var row = range.getRow();

    // Check if the edit is in the ACTION_COL and is the "Send Email" text
    if (column === CONFIG.SHEET_VALIDATION.ACTION_COL && row > 1) { 
      var action = range.getValue();
      if (action === CONFIG.DROPDOWN_ACTION_TEXT) {
        sendEmailForRow(row, sheet);
        range.setValue(CONFIG.DROPDOWN_SENT_TEXT); // Update to "Email Sent"
      }
    }
  }
}

/**
 * Function to send reminder emails in bulk from "Reminder Tool" sheet.
 */
function sendReminderEmailsInBulk() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_REMINDER.NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Error: Sheet '" + CONFIG.SHEET_REMINDER.NAME + "' not found.");
    return;
  }

  var startRow = CONFIG.SHEET_REMINDER.START_ROW;
  var lastRow = sheet.getLastRow();

  if (lastRow < startRow) {
    SpreadsheetApp.getUi().alert("No email data found to process.");
    return;
  }

  // Get all data in one go
  var range = sheet.getRange(startRow, 1, lastRow - startRow + 1, CONFIG.SHEET_REMINDER.NAME_COL);
  var data = range.getValues();

  for (var i = 0; i < data.length; i++) {
    var currentRow = startRow + i;
    var recipientEmail = data[i][CONFIG.SHEET_REMINDER.EMAIL_COL - 1] ? data[i][CONFIG.SHEET_REMINDER.EMAIL_COL - 1].toString().trim() : "";
    var recipientName = data[i][CONFIG.SHEET_REMINDER.NAME_COL - 1] ? data[i][CONFIG.SHEET_REMINDER.NAME_COL - 1].toString().trim() : "";

    if (!recipientEmail || !recipientName) {
      sheet.getRange(currentRow, CONFIG.SHEET_REMINDER.STATUS_COL).setValue("Skipped: Missing email or name");
      continue;
    }

    sendEmail(recipientEmail, null, recipientName, sheet, currentRow, "Reminder Tool");
  }

  SpreadsheetApp.getUi().alert("Bulk email sending process completed.");
}

/**
 * Reusable function to send email for both tools.
 * @param {string} toEmail - The primary recipient's email.
 * @param {string|null} ccEmail - The CC email (for Validation) or null (for Reminder).
 * @param {string} recipientName - The name of the recipient (e.g., "Company Name").
 * @param {Sheet} sheet - The Google Sheet object.
 * @param {number} row - The row number being processed.
 * @param {string} sheetType - The name of the sheet ("Validation" or "Reminder Tool").
 */
function sendEmail(toEmail, ccEmail, recipientName, sheet, row, sheetType) {
  
  // Build subject from template
  var subject = CONFIG.EMAIL_SUBJECT_TEMPLATE.replace("${recipientName}", recipientName);

  // 1. Get the HTML template from the file
  var htmlTemplate;
  try {
     htmlTemplate = HtmlService.createTemplateFromFile("emailTemplate.html");
  } catch (e) {
    SpreadsheetApp.getUi().alert("FATAL ERROR: HTML template file 'emailTemplate.html' not found. Please create it.");
    return;
  }
  
  // 2. Set the variables for the template
  htmlTemplate.recipientName = recipientName;
  // You can add more variables here, e.g.:
  // htmlTemplate.someOtherData = sheet.getRange(row, 7).getValue(); // Get data from Col G

  // 3. "Evaluate" the template to build the final HTML string
  var emailBody = htmlTemplate.evaluate().getContent();

  // 4. Set up email options
  var emailOptions = {
    htmlBody: emailBody,
    subject: subject,
    from: CONFIG.SENDER_EMAIL,
    name: CONFIG.SENDER_NAME
  };

  try {
    if (sheetType === "Validation") {
      emailOptions.to = ccEmail; // This is the 'requestorEmail'
      emailOptions.cc = toEmail; // This is the 'recipientEmail'
      
      MailApp.sendEmail(emailOptions);
      
      sheet.getRange(row, CONFIG.SHEET_VALIDATION.STATUS_COL).setValue("Sent");
    
    } else { // Reminder Tool
      emailOptions.to = toEmail;
      
      MailApp.sendEmail(emailOptions);
      
      var timestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "MM/dd/yyyy HH:mm:ss");
      sheet.getRange(row, CONFIG.SHEET_REMINDER.TIMESTAMP_COL).setValue(timestamp);
      sheet.getRange(row, CONFIG.SHEET_REMINDER.STATUS_COL).setValue("Sent");
    }
  } catch (error) {
    var errorMsg = "Error: " + error.message;
    if (sheetType === "Validation") {
      sheet.getRange(row, CONFIG.SHEET_VALIDATION.STATUS_COL).setValue(errorMsg);
    } else {
      sheet.getRange(row, CONFIG.SHEET_REMINDER.STATUS_COL).setValue(errorMsg);
    }
  }
}

/**
 * Function to send email for a specific row in "Validation" sheet.
 */
function sendEmailForRow(row, sheet) {
  var requestorEmail = sheet.getRange(row, CONFIG.SHEET_VALIDATION.SENDER_COL).getValue().toString().trim();
  var recipientName = sheet.getRange(row, CONFIG.SHEET_VALIDATION.RECIPIENT_NAME_COL).getValue().toString().trim();
  var recipientEmail = sheet.getRange(row, CONFIG.SHEET_VALIDATION.RECIPIENT_EMAIL_COL).getValue().toString().trim();

  if (!requestorEmail || !recipientEmail || !recipientName) {
    sheet.getRange(row, CONFIG.SHEET_VALIDATION.STATUS_COL).setValue("Skipped: Missing data");
    return;
  }

  // Note: For the Validation sheet, the 'to' is the requestor and the 'cc' is the recipient
  // We pass 'recipientEmail' as 'toEmail' and 'requestorEmail' as 'ccEmail'
  sendEmail(recipientEmail, requestorEmail, recipientName, sheet, row, "Validation");
}
