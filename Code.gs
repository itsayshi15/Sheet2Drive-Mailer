const CONFIG = {
  SHEET_NAME: "Sheet1",                          // Exact name of your sheet tab
  DRAFT_SUBJECT: "Your Certificate",             // Exact subject of your Gmail draft
  SENDER_NAME: "Your Sender Name Here",          // <-- change this
  CERTIFICATES_FOLDER_ID: "1ZWXmpgWo8hG2_dPUzQNvbSVuP8K-kqB4",
  EXTRA_ATTACHMENT_IDS: [],                      // Optional files sent to everyone

  // ALWAYS added to every email, in addition to whatever the draft has.
  // Example: "a@gmail.com, b@gmail.com". Leave "" if not needed.
  // This is the 100% guaranteed way to get Cc/Bcc on every email.
  FALLBACK_CC: "",
  FALLBACK_BCC: "",

  SEND_DELAY_MS: 300,
  MAX_RUNTIME_MS: 5 * 60 * 1000,                 // Stop before Apps Script's 6 min limit
};

const HANDLER = "onChangeSendEmail";
let _pdfMap = null;

/* ---------- Menu (manual, instant run) ---------- */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Certificates")
    .addItem("Send pending now", "onChangeSendEmail")
    .addItem("Test draft Cc/Bcc", "testDraftCcBcc")
    .addToUi();
}

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

/** Shows on screen what Cc/Bcc the script finds. Sends nothing. */
function testDraftCcBcc() {
  const draft = findDraft(CONFIG.DRAFT_SUBJECT);
  let text;
  if (!draft) {
    text = `Draft with subject "${CONFIG.DRAFT_SUBJECT}" NOT FOUND.\nCheck DRAFT_SUBJECT matches exactly.`;
  } else {
    const r = getDraftCcBcc(draft);
    text = `FINAL Cc: ${r.cc || "(none)"}\nFINAL Bcc: ${r.bcc || "(none)"}\n\n--- Details ---\n${r.debug.join("\n")}`;
  }
  Logger.log(text);
  try {
    SpreadsheetApp.getUi().alert("Draft Cc/Bcc check", text, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (err) {
    // Not running from the sheet menu; log only.
  }
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

    const ccbcc = getDraftCcBcc(draftToSend);
    const draftCc = ccbcc.cc;
    const draftBcc = ccbcc.bcc;
    Logger.log(`Using Cc="${draftCc}" Bcc="${draftBcc}"\n` + ccbcc.debug.join("\n"));

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
        setStatus(sheet, rowNum, "Failed: name is empty");
        continue;
      }

      const emailList = email.split(/[,;]/)
        .map(a => a.trim())
        .filter(a => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(a));

      if (emailList.length === 0) {
        setStatus(sheet, rowNum, "Failed: no valid email");
        continue;
      }

      if (MailApp.getRemainingDailyQuota() < emailList.length) {
        Logger.log("Daily email quota exhausted - stopping.");
        break;
      }

      const customPdf = getCustomCertificate(name);
      if (!customPdf) {
        setStatus(sheet, rowNum, `Failed: PDF missing for '${name}.pdf'`);
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

/** Writes status and pushes it to the sheet immediately. */
function setStatus(sheet, row, value) {
  sheet.getRange(row, 3).setValue(value);
  SpreadsheetApp.flush();
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

/** Finds the draft by subject. If several match, uses the most recent one. */
function findDraft(subject) {
  const target = subject.trim().toLowerCase();
  const drafts = GmailApp.getDrafts();
  let best = null;
  let bestTime = 0;
  let matches = 0;

  for (let i = 0; i < drafts.length; i++) {
    const msg = drafts[i].getMessage();
    if (msg.getSubject().trim().toLowerCase() === target) {
      matches++;
      const t = msg.getDate().getTime();
      if (!best || t > bestTime) {
        best = drafts[i];
        bestTime = t;
      }
    }
  }
  if (matches > 1) {
    Logger.log(`Warning: ${matches} drafts share this subject. Using the most recently edited one.`);
  }
  return best;
}

/** Extracts plain email addresses from any header-like string. */
function extractEmails(str) {
  if (!str) return [];
  const found = str.toString().match(/[A-Za-z0-9._%+\-']+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/g);
  return found || [];
}

/** Reads a header (e.g. "Cc") from a raw RFC822 message, handling folded lines. */
function getRawHeader(raw, headerName) {
  const headerPart = raw.split(/\r?\n\r?\n/)[0];
  const unfolded = headerPart.replace(/\r?\n[ \t]+/g, " ");
  const re = new RegExp("^" + headerName + ":\\s*(.*)$", "im");
  const m = unfolded.match(re);
  return m ? m[1] : "";
}

/** Merges lists of addresses, removing duplicates. */
function mergeEmails() {
  const seen = {};
  const out = [];
  for (let i = 0; i < arguments.length; i++) {
    extractEmails(arguments[i]).forEach(a => {
      const k = a.toLowerCase();
      if (!seen[k]) { seen[k] = true; out.push(a); }
    });
  }
  return out.join(",");
}

/**
 * Tries 4 ways to read Cc/Bcc from the draft, then adds the CONFIG fallbacks.
 * Returns {cc, bcc, debug[]}.
 */
function getDraftCcBcc(draft) {
  const message = draft.getMessage();
  const debug = [];
  let cc = "", bcc = "";

  // Method 1: GmailApp built-in
  try {
    const c = message.getCc() || "";
    const b = message.getBcc() || "";
    cc = mergeEmails(cc, c);
    bcc = mergeEmails(bcc, b);
    debug.push(`1 GmailApp: Cc="${c}" Bcc="${b}"`);
  } catch (err) {
    debug.push("1 GmailApp FAILED: " + err.message);
  }

  // Method 2: raw message headers
  try {
    const raw = message.getRawContent();
    const c = getRawHeader(raw, "Cc");
    const b = getRawHeader(raw, "Bcc");
    cc = mergeEmails(cc, c);
    bcc = mergeEmails(bcc, b);
    debug.push(`2 Raw headers: Cc="${c}" Bcc="${b}"`);
  } catch (err) {
    debug.push("2 Raw headers FAILED: " + err.message);
  }

  // Method 3: Advanced Gmail Service (only if enabled under Services)
  try {
    if (typeof Gmail === "undefined") throw new Error("Gmail service not enabled (Services > Gmail API)");
    const d = Gmail.Users.Drafts.get("me", draft.getId(), { format: "full" });
    const headers = (d.message && d.message.payload && d.message.payload.headers) || [];
    const findHeader = name => {
      const h = headers.find(h => h.name.toLowerCase() === name);
      return h ? h.value : "";
    };
    const c = findHeader("cc");
    const b = findHeader("bcc");
    cc = mergeEmails(cc, c);
    bcc = mergeEmails(bcc, b);
    debug.push(`3 Gmail service: Cc="${c}" Bcc="${b}"`);
  } catch (err) {
    debug.push("3 Gmail service skipped: " + err.message);
  }

  // Method 4: Gmail REST API using the script's own token (no service needed)
  if (!cc && !bcc) {
    try {
      const url = "https://gmail.googleapis.com/gmail/v1/users/me/drafts/" + draft.getId() +
        "?format=metadata&metadataHeaders=Cc&metadataHeaders=Bcc";
      const res = UrlFetchApp.fetch(url, {
        headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true,
      });
      if (res.getResponseCode() === 200) {
        const d = JSON.parse(res.getContentText());
        const headers = (d.message && d.message.payload && d.message.payload.headers) || [];
        const findHeader = name => {
          const h = headers.find(h => h.name.toLowerCase() === name);
          return h ? h.value : "";
        };
        const c = findHeader("cc");
        const b = findHeader("bcc");
        cc = mergeEmails(cc, c);
        bcc = mergeEmails(bcc, b);
        debug.push(`4 REST API: Cc="${c}" Bcc="${b}"`);
      } else {
        debug.push("4 REST API skipped: HTTP " + res.getResponseCode());
      }
    } catch (err) {
      debug.push("4 REST API skipped: " + err.message);
    }
  }

  // Add guaranteed addresses from CONFIG
  cc = mergeEmails(cc, CONFIG.FALLBACK_CC);
  bcc = mergeEmails(bcc, CONFIG.FALLBACK_BCC);

  if (!cc && !bcc) {
    debug.push("WARNING: no Cc/Bcc found anywhere. Set FALLBACK_CC / FALLBACK_BCC in CONFIG.");
  }

  return { cc: cc, bcc: bcc, debug: debug };
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
      Logger.log(`Sent certificate to ${name} <${recipient}> | Cc="${draftCc}" Bcc="${draftBcc}"`);
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

  setStatus(sheet, row, statusValue);
}

/* ---------- Sheet setup ---------- */

function setupHeaders(sheet) {
  if (sheet.getLastRow() === 0 || sheet.getRange("A1").getValue() === "") {
    sheet.getRange("A1:C1").setValues([["Email", "Name", "Status"]]);
    sheet.getRange("A1:C1").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
}
