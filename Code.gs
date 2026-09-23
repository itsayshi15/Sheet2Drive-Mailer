const CONFIG = {
  SHEET_NAME: "Sheet1",                          // Exact name of your sheet tab
  DRAFT_SUBJECT: "Your Draft Subject Line Here", // Exact subject of your Gmail draft
  SENDER_NAME: "Your Sender Name Here",
  CERTIFICATES_FOLDER_ID: "YOUR_GOOGLE_DRIVE_FOLDER_ID_HERE",
  EXTRA_ATTACHMENT_IDS: [],                      // Optional files sent to everyone
  SEND_DELAY_MS: 300,
  MAX_RUNTIME_MS: 5 * 60 * 1000,                 // Stop before Apps Script's 6 min limit
};

const HANDLER = "onChangeSendEmail";
let _pdfMap = null;

/** Run this ONCE manually to install the trigger. */
function createTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === HANDLER)
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger(HANDLER)
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onChange()
    .create();

  Logger.log("Trigger installed. Paste rows and emails auto-send.");
}

function onChangeSendEmail(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log("Another execution is running - skipping this run.");
    return;
  }

  const startTime = Date.now();

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) {
      Logger.log(`Sheet "${CONFIG.SHEET_NAME}" not found.`);
      return;
    }
    setupHeaders(sheet);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

    // Only continue if there is something to process
    const hasWork = data.some(r => {
      const email = (r[0] || "").toString().trim();
      const status = (r[2] || "").toString().trim();
      return email && (status === "" || /^failed/i.test(status));
    });
    if (!hasWork) return;

    const draftToSend = findDraft(CONFIG.DRAFT_SUBJECT);
    if (!draftToSend) {
      Logger.log("Draft not found - nothing sent this run.");
      return;
    }
    const message = draftToSend.getMessage();
    const { cc: draftCc, bcc: draftBcc } = getDraftCcBcc(draftToSend.getId());

    const extraAttachments = CONFIG.EXTRA_ATTACHMENT_IDS.map(id => {
      try {
        return DriveApp.getFileById(id).getBlob();
      } catch (err) {
        Logger.log(`Could not load extra attachment ${id}: ${err.message}`);
        return null;
      }
    }).filter(Boolean);

    const alreadySent = getAlreadySent(data);
    _pdfMap = null; // reload PDF list fresh each run

    for (let idx = 0; idx < data.length; idx++) {
      if (Date.now() - startTime > CONFIG.MAX_RUNTIME_MS) {
        Logger.log("Time limit near - stopping. Remaining rows will be sent on the next change/run.");
        break;
      }

      const rowNum = idx + 2;
      const email = (data[idx][0] || "").toString().trim();
      const name = (data[idx][1] || "").toString().trim();
      const status = (data[idx][2] || "").toString().trim();

      // Process empty status, or retry rows that previously failed
      if (!email) continue;
      if (status !== "" && !/^failed/i.test(status)) continue;

      if (!name) {
        sheet.getRange(rowNum, 3).setValue("Failed: name is empty");
        continue;
      }

      const emailList = email.split(/[,;]/)
        .map(a => a.trim())
        .filter(a => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(a));

      if (emailList.length === 0) {
        sheet.getRange(rowNum, 3).setValue("Failed: no valid email");
        continue;
      }

      if (MailApp.getRemainingDailyQuota() < emailList.length) {
        Logger.log("Daily email quota exhausted - stopping.");
        break;
      }

      const customPdf = getCustomCertificate(name);
      if (!customPdf) {
        sheet.getRange(rowNum, 3).setValue(`Failed: PDF missing for '${name}.pdf'`);
        Logger.log(`Skipped ${name}: certificate not found in Drive folder.`);
        continue;
      }

      sendRowEmails(sheet, rowNum, emailList, name, message, customPdf, extraAttachments, alreadySent, draftCc, draftBcc);
      Utilities.sleep(CONFIG.SEND_DELAY_MS);
    }
  } finally {
    lock.releaseLock();
  }
}

/* ---------- PDF lookup (case/space-insensitive) ---------- */

function normalizeName(s) {
  return s.toString().replace(/\.pdf$/i, "").replace(/\s+/g, " ").trim().toLowerCase();
}

function loadPdfMap() {
  _pdfMap = {};
  if (!CONFIG.CERTIFICATES_FOLDER_ID || CONFIG.CERTIFICATES_FOLDER_ID === "YOUR_GOOGLE_DRIVE_FOLDER_ID_HERE") {
    Logger.log("CERTIFICATES_FOLDER_ID is missing in CONFIG.");
    return;
  }
  const files = DriveApp.getFolderById(CONFIG.CERTIFICATES_FOLDER_ID).getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType() === MimeType.PDF) {
      const key = normalizeName(f.getName());
      if (!_pdfMap[key]) _pdfMap[key] = f; // first match wins
    }
  }
  Logger.log(`Loaded ${Object.keys(_pdfMap).length} PDFs from folder.`);
}

function getCustomCertificate(recipientName) {
  try {
    if (!_pdfMap) loadPdfMap();
    const file = _pdfMap[normalizeName(recipientName)];
    return file ? file.getBlob() : null;
  } catch (err) {
    Logger.log(`Error finding certificate for ${recipientName}: ${err.message}`);
    return null;
  }
}

/* ---------- Draft helpers ---------- */

function findDraft(subject) {
  const target = subject.trim().toLowerCase();
  const drafts = GmailApp.getDrafts();
  for (let i = 0; i < drafts.length; i++) {
    if (drafts[i].getMessage().getSubject().trim().toLowerCase() === target) {
      return drafts[i];
    }
  }
  return null;
}

function getDraftCcBcc(draftId) {
  try {
    const draft = Gmail.Users.Drafts.get("me", draftId, { format: "full" });
    const headers = (draft.message && draft.message.payload && draft.message.payload.headers) || [];
    const findHeader = name => {
      const h = headers.find(h => h.name.toLowerCase() === name);
      return h ? h.value : "";
    };
    return { cc: findHeader("cc"), bcc: findHeader("bcc") };
  } catch (err) {
    Logger.log("Advanced Gmail API not enabled (Services > Gmail API). Cc/Bcc ignored.");
    return { cc: "", bcc: "" };
  }
}

function fillPlaceholders(text, name) {
  return text.replace(/{{\s*name\s*}}/gi, name);
}

/* ---------- Duplicate tracking (email + name) ---------- */

function sentKey(email, name) {
  return email.trim().toLowerCase() + "|" + normalizeName(name);
}

function getAlreadySent(data) {
  const sent = new Set();
  data.forEach(row => {
    const name = (row[1] || "").toString();
    const status = (row[2] || "").toString();

    if (status === "Sent") {
      (row[0] || "").toString().split(/[,;]/)
        .map(e => e.trim()).filter(Boolean)
        .forEach(e => sent.add(sentKey(e, name)));
    } else if (/^partial/i.test(status)) {
      const match = status.match(/sent to (.*?)(?:;\s*(?:duplicate|failed)|$)/i);
      if (match) {
        match[1].split(",")
          .map(e => e.trim()).filter(Boolean)
          .forEach(e => sent.add(sentKey(e, name)));
      }
    }
  });
  return sent;
}

/* ---------- Sending ---------- */

function sendRowEmails(sheet, row, emailList, name, message, customPdf, extraAttachments, alreadySent, draftCc, draftBcc) {
  const attachments = [customPdf].concat(message.getAttachments()).concat(extraAttachments || []);

  const subject = fillPlaceholders(message.getSubject(), name);
  const plainBody = fillPlaceholders(message.getPlainBody(), name);
  const htmlBody = fillPlaceholders(message.getBody(), name);

  const succeeded = [];
  const failed = [];
  const duplicates = [];

  emailList.forEach(recipient => {
    const key = sentKey(recipient, name);

    if (alreadySent.has(key)) {
      duplicates.push(recipient);
      return;
    }

    try {
      const mailOptions = {
        htmlBody: htmlBody,
        attachments: attachments,
        name: CONFIG.SENDER_NAME,
      };
      if (draftCc) mailOptions.cc = draftCc;
      if (draftBcc) mailOptions.bcc = draftBcc;

      GmailApp.sendEmail(recipient, subject, plainBody, mailOptions);
      succeeded.push(recipient);
      alreadySent.add(key);
      Logger.log(`Sent certificate to ${name} <${recipient}>`);
    } catch (err) {
      failed.push(recipient + " (" + err.message + ")");
      Logger.log(`Failed for ${name} <${recipient}>: ` + err.message);
    }
  });

  const parts = [];
  if (succeeded.length) parts.push("sent to " + succeeded.join(", "));
  if (duplicates.length) parts.push("duplicate (already sent): " + duplicates.join(", "));
  if (failed.length) parts.push("failed: " + failed.join("; "));

  let statusValue;
  if (succeeded.length === 0 && failed.length === 0 && duplicates.length > 0) {
    statusValue = "Duplicate - already sent: " + duplicates.join(", ");
  } else if (failed.length === 0 && duplicates.length === 0) {
    statusValue = "Sent";
  } else if (succeeded.length === 0 && failed.length > 0) {
    statusValue = "Failed: " + failed.join("; ");
  } else {
    statusValue = "Partial - " + parts.join("; ");
  }

  sheet.getRange(row, 3).setValue(statusValue);
}

/* ---------- Sheet setup ---------- */

function setupHeaders(sheet) {
  if (sheet.getLastRow() === 0 || sheet.getRange("A1").getValue() === "") {
    sheet.getRange("A1:C1").setValues([["Email", "Name", "Status"]]);
    sheet.getRange("A1:C1").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
} 
