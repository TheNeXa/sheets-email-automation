# Google Sheets Email Automation

This Google Apps Script provides a robust framework for sending data-driven, templated emails directly from Google Sheets. It supports two distinct workflows:
1.  **Bulk Sender:** Sends mass emails to a list of recipients defined in one sheet.
2.  **Ad-Hoc Trigger:** Uses an interactive dropdown menu in another sheet to send individual emails, perfect for task-based reminders.

## Features

* **Dual Workflows:** Manage both bulk email campaigns and individual, row-based email triggers from one script.
* **Centralized Configuration:** Easily update all settings (sheet names, column numbers, email subjects) in a single `CONFIG` object.
* **HTML Templating:** Uses an `emailTemplate.html` file to send rich, dynamic HTML emails.
* **Status Tracking:** Automatically updates the sheet with "Sent" or "Error" statuses and timestamps.
* **Custom Menu:** Adds a simple menu (`Email Tool`) to the Google Sheet UI to run primary functions.
* **Conditional Formatting:** Visually highlights the status of ad-hoc triggers (e.g., "Send Email" vs. "Email Sent").

## Repository Structure
├── Code.gs             # The main Google Apps Script file
├── emailTemplate.html  # The HTML template for outgoing emails
└── README.md           # This file

## Installation

1.  **Create Google Sheet:**
    * Create a new Google Sheet.
    * Create two tabs. Name them **"Bulk Tasks"** and **"Ad-Hoc Tasks"** (or change the names in the `CONFIG` object in `Code.gs`).
2.  **Open Apps Script:**
    * In your Google Sheet, go to `Extensions > Apps Script`.
3.  **Add Script Files:**
    * Copy the contents of `Code.gs` from this repository and paste it into the `Code.gs` file in the Apps Script editor.
    * Click the **+** icon and select **HTML**. Name the new file **`emailTemplate.html`**.
    * Copy the contents of `emailTemplate.html` from this repository and paste it into the new file.
4.  **Save and Authorize:**
    * Save the project.
    * Reload your Google Sheet. A new menu named **"Email Tool"** should appear.
    * You will need to authorize the script the first time you run a function.

## Configuration

All settings are managed in the `CONFIG` object at the top of `Code.gs`.

```javascript
const CONFIG = {
  // General settings: sender name, timezone, email subject
  SENDER_NAME: "Your Team Name",
  TIMEZONE: "Asia/Jakarta",
  EMAIL_SUBJECT_TEMPLATE: "Action Required: Follow-up for ${recipientName}",

  // Sheet for Ad-Hoc (dropdown) tasks
  SHEET_ADHOC: {
    NAME: "Ad-Hoc Tasks",
    SENDER_COL: 2,           // B: Requestor Email (Internal)
    RECIPIENT_NAME_COL: 3,   // C: Recipient Name
    RECIPIPIENT_EMAIL_COL: 6,  // F: Recipient Email (External)
    ACTION_COL: 15,          // O: Action Dropdown
    STATUS_COL: 16           // P: Email Status
  },
  
  // Sheet for Bulk email tasks
  SHEET_BULK: {
    NAME: "Bulk Tasks",
    EMAIL_COL: 1,            // A: Email Address
    NAME_COL: 2,             // B: Recipient Name
    TIMESTAMP_COL: 3,        // C: Last Send History
    STATUS_COL: 4,           // D: Send Status
    START_ROW: 4             // The first row containing data
  },
  
  // Settings for the dropdown formatting
  DROPDOWN_OPTIONS: ["Send Email", "Email Sent"],
  // ...other formatting settings
};
```
## Usage

### 1. Ad-Hoc Email Trigger (The "Ad-Hoc Tasks" Sheet)

This sheet is for sending individual, task-based emails.

1.  **Run Setup:** From the Google Sheet menu, select **Email Tool > Setup Ad-Hoc Triggers**. This will add the headers, dropdowns, and formatting to the `Ad-Hoc Tasks` sheet.
2.  **Populate Data:** Add your data to the sheet starting from row 2:
    * **Column B (Sender):** The *internal* email of the person making the request.
    * **Column C (Recipient Name):** The name of the external recipient.
    * **Column F (Recipient Email):** The email address of the external recipient.
3.  **Send Email:** Go to **Column O (Action)** and select **"Send Email"** from the dropdown.

> **Workflow Note:** This trigger sends an email **TO** the "Sender" (Col B) and **CCs** the "Recipient" (Col F). This is designed for internal tracking, where a team member is notified that a reminder has been sent *to* the recipient.

### 2. Bulk Email Sender (The "Bulk Tasks" Sheet)

This sheet is for sending a mass email to a list.

1.  **Populate Data:** Add your data to the sheet starting from **row 4** (or as defined in `CONFIG.SHEET_BULK.START_ROW`):
    * **Column A (Email Address):** The recipient's email.
    * **Column B (Recipient Name):** The recipient's name.
2.  **Send Emails:** From the Google Sheet menu, select **Email Tool > Send Bulk Emails**.

The script will loop through all rows, send the emails, and update the "Last Send History" (Col C) and "Send Status" (Col D) for each row.

### 3. Customizing the Email

To change the email content, edit the **`emailTemplate.html`** file. You can use standard HTML/CSS and insert variables from your script.

* The variable `<?!= recipientName ?>` is used by default.
* You can add more variables (e.g., `<?!= someOtherData ?>`) by modifying the `sendEmail` function in `Code.gs`.
