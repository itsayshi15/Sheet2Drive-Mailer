# Sheet2Drive-Mailer: Google Sheets & Drive Auto Document Mailer

**Sheet2Drive-Mailer** is an automated Google Apps Script solution that monitors a Google Sheet, matches recipient names with PDF document stored in a Google Drive folder, and dispatches personalized emails using a Gmail draft template.

---

## Key Features

* **Automatic Drive Matching:** Scans a specified Google Drive folder to match PDF filenames (e.g., `Jane Doe.pdf`) with names in your spreadsheet.
* **Event-Driven Trigger:** Automatically triggers when new rows or data updates occur (`onChange` trigger).
* **Draft Template Integration:** Uses a Gmail draft as a template with dynamic `{{name}}` tags and inherited HTML formatting.
* **Concurrency Protection:** Employs `LockService` to prevent race conditions during simultaneous sheet edits.
* **Status & Retry Tracking:** Logs exact delivery statuses (`Sent`, `Failed`, `Partial`, or `Duplicate`) directly back to the sheet.
* **Quota & Timeout Safe:** Respects Gmail daily limits and Apps Script's 6-minute execution window.

---

## How It Works

1. **Sheet Setup**: Enter recipient data into your sheet:
   * **Column A**: `Email`
   * **Column B**: `Name`
   * **Column C**: `Status`

2. **Drive Lookup**: CertiMail searches your designated Google Drive folder for `<Name>.pdf` (case and space-insensitive).

3. **Email Dispatch**: It personalizes your Gmail draft template, attaches the matching PDF certificate from Drive, and sends it to the recipient's email.

---

## Setup & Installation

### Step 1: Add Script to Google Sheets
1. Open your Google Sheet.
2. Go to **Extensions** > **Apps Script**.
3. Copy and paste `Code.gs` into the editor.

### Step 2: Configure Parameters
Update the `CONFIG` object at the top of the file:

```javascript
const CONFIG = {
  SHEET_NAME: "Sheet1",                  // Tab name in Google Sheets
  DRAFT_SUBJECT: "Your Certificate is Ready!",   // Subject line of your saved Gmail draft
  SENDER_NAME: "Event Organizer",                // Sender name displayed in recipient inboxes
  CERTIFICATES_FOLDER_ID: "YOUR_FOLDER_ID_HERE", // Google Drive folder ID containing certificate PDFs
  EXTRA_ATTACHMENT_IDS: [],                      // (Optional) Universal file IDs sent to everyone
  SEND_DELAY_MS: 300,                            // Delay between sends (in ms)
  MAX_RUNTIME_MS: 5 * 60 * 1000,                  // 5-minute safety threshold
};
```

### Step 3: Initialize Trigger

1. Run `createTrigger()` once manually inside the Apps Script editor.
2. Accept the required authorization prompts.

---

## Spreadsheet Layout

| Email | Name | Status |
| --- | --- | --- |
| `jane.doe@example.com` | `Jane Doe` | *(Updated automatically)* |
| `john.smith@example.com` | `John Smith` | *(Updated automatically)* |

---

## File Naming Convention

Certificate files in your Google Drive folder must match the names listed in **Column B**:

* **Sheet Name**: `Jane Doe`
* **Drive File**: `Jane Doe.pdf` *(or `jane doe.pdf`)*

---

## Optional: CC / BCC Support

To preserve `CC` and `BCC` fields from your Gmail draft template:

1. Open Apps Script editor.
2. Click **Services** (`+`) in the left panel.
3. Select and add **Gmail API**.


