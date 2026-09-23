
# Smart Drive Mail Merge: Google Sheets & Drive Auto Document Mailer

An automated Google Apps Script solution that sends personalized documents (example: certificates) and email notifications using Google Sheets, Google Drive, and Gmail draft templates.

---

##  Why This Project?

Manual distribution of certificates after events, webinars, or courses can be tedious, error-prone, and time-consuming. 

This project automates the entire distribution pipeline directly inside Google Workspace:
* **Zero External Dependencies:** Runs natively within Google Workspace using built-in Google Apps Script—no third-party paid tools required.
* **Template Flexibility:** Design your email directly inside Gmail as a draft. The script retains all formatting, inline images (logos, signatures), and CC/BCC configurations.
* **Fail-Safe & Smart Tracking:** Built-in safeguards prevent duplicate emails, handle batch timeouts gracefully, and track delivery status row-by-row.

---

##  Features

*  **Personalized Email Templates:** Replaces `{{name}}` placeholders in the Gmail draft subject and body with recipient names.
* **Dynamic Documents Matching:** Automatically pairs recipient names from the sheet with matching PDF files in your Google Drive folder (`Recipient Name.pdf`).
* **Inline Images & Attachments:** Preserves embedded logos, signatures, and extra attachments attached to the Gmail draft template.
* **Advanced CC/BCC Preservation:** Automatically detects CC and BCC recipients set on the Gmail draft.
* **Smart Status & Retry System:**
  * Auto-detects pending rows.
  * Tracks statuses (`Sent`, `Failed`, `Partial`, `Duplicate`).
  * Allows force resending by setting status to `Resend`.
* **Safety Limits & Quota Protections:**
  * Uses **Script Lock** to prevent concurrent executions.
  * Monitors daily Google Mail quota to avoid quota exceeding errors.
  * Respects the 6-minute Google Apps Script execution timeout limit.
* **Custom UI Menu:** Adds a custom menu directly inside Google Sheets for quick execution and CC/BCC testing.

---

##  Setup & Usage Instructions

### 1. Google Drive Preparation
1. Create a folder in Google Drive and upload all generated PDF certificates.
2. Ensure PDF file names match the exact recipient names used in your Google Sheet (e.g., `John Doe.pdf`).
3. Copy the **Folder ID** from the Google Drive URL:
   `https://drive.google.com/drive/folders/YOUR_FOLDER_ID_HERE`

### 2. Gmail Draft Preparation
1. Open Gmail and create a new draft email.
2. Set the subject line (e.g., `Your Certificate of Completion`).
3. Compose your body text. Use `{{name}}` wherever you want the recipient's name to appear.
4. Add any signatures, inline logos, CC/BCC addresses, or additional attachments as needed.
5. Save the draft (**do not send it**).

### 3. Google Sheet Setup
1. Create or open your Google Sheet.
2. Ensure your active tab name matches `CONFIG.SHEET_NAME` (default is `Sheet1`).
3. Set up the column headers in Row 1:
   * **Column A:** `Email`
   * **Column B:** `Name`
   * **Column C:** `Status`

### 4. Install the Script
1. In your Google Sheet, open **Extensions > Apps Script**.
2. Replace all code in `Code.gs` with the project script code.
3. Update the `CONFIG` object at the top of the file:

```javascript
const CONFIG = {
  SHEET_NAME: "Sheet1",                 // Name of your sheet tab
  DRAFT_SUBJECT: "Your Subject Here",   // Subject line of your Gmail draft
  SENDER_NAME: "Your Name / Org Name",  // Sender display name
  CERTIFICATES_FOLDER_ID: "FOLDER_ID",  // Google Drive folder ID containing PDFs
  EXTRA_ATTACHMENT_IDS: [],             // Optional extra file IDs sent to everyone
  FALLBACK_CC: "",                      // Optional fallback CC email address
  FALLBACK_BCC: "",                     // Optional fallback BCC email address
  SEND_DELAY_MS: 300,                   // Delay between emails in milliseconds
  MAX_RUNTIME_MS: 5 * 60 * 1000,        // Timeout protection (5 minutes)
};
```

4. Save the project (`Ctrl + S` or `Cmd + S`).

### 5. Authorize & Initialize

1. **Test Draft Settings:**
   * Select `testDraftCcBcc` from the function dropdown at the top of Apps Script and click **Run**.
   * Review the execution log to confirm your Gmail draft's CC/BCC settings are properly detected.

2. **Install Automatic Trigger:**
   * Select `createTrigger` from the function dropdown and click **Run**.
   * Authorize the necessary permissions when prompted by Google.
   * *This creates an `onChange` trigger that automatically processes new rows added to the sheet.*

3. **Manual Run / Sheet Menu:**
   * Refresh your Google Sheet.
   * Use the custom menu: **Certificates > Send pending now** to run the execution manually at any time.

---

## Status Reference Column

| Status | Behavior |
| :--- | :--- |
| *(Blank)* | Marked as **Pending**. Processed on the next run. |
| `Resend` | Forces a re-send to the recipient, bypassing duplicate detection. |
| `Failed...` | Marked for retry on subsequent runs. |
| `Sent` | Successfully sent. Will be skipped on future runs. |
| `Duplicate of row N` | Skipped because this recipient/name pair already received an email. |
