const CONFIG = {
  SHEET_NAME: "Sheet1",                          // Exact name of your sheet tab
  DRAFT_SUBJECT: "XXXXXXXXXX",             // Exact subject of your Gmail draft
  SENDER_NAME: "XXXXXXXXX",          // <-- change this
  CERTIFICATES_FOLDER_ID: "XXXXXXXXX",
  EXTRA_ATTACHMENT_IDS: [],                      // Optional files sent to everyone

  // ALWAYS added to every email, in addition to whatever the draft has.
  FALLBACK_CC: "",
  FALLBACK_BCC: "",

  SEND_DELAY_MS: 300,
  MAX_RUNTIME_MS: 5 * 60 * 1000,                 // Stop before Apps Script's 6 min limit
};

const HANDLER = "onChangeSendEmail";
let _pdfMap = null;

/* ---------- Menu ---------- */

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

/* ---------- Status rules ----------
 * ""            -> pending (new row)
 * "Failed..."   -> pending (retried)
 * "Resend"      -> pending, and skips duplicate check (use to force a re-send)
 * "Sent", "Partial...", "Sending...", "Duplicate of row N" -> never processed
 */

function isPending(status) {
  return status === "" || /^failed/i.test(status) || /^resend/i.test(status) || /^not sent/i.test(status);
}

/**
 * Marks every pending row from startRow down as "Not sent: <reason>".
 * Reads fresh statuses from the sheet, so rows already handled this run are untouched.
 * Rows that were "Resend" keep that intent as "Not sent (resend): ...".
 */
function markNotSent(sheet, reason, startRow) {
  const last = sheet.getLastRow();
  const from = Math.max(startRow || 2, 2);
  const n = last - from + 1;
  if (n < 1) return;

  const vals = sheet.getRange(from, 1, n, 3).getValues();
  const out = vals.map(r => {
    const email = (r[0] || "").toString().trim();
    const status = (r[2] || "").toString().trim();
    if (!email || !isPending(status)) return [r[2]];
    const tag = /^(resend|not sent \(resend\))/i.test(status) ? "Not sent (resend)" : "Not sent";
    return [`${tag}: ${reason}`];
  });
  sheet.getRange(from, 3, n, 1).setValues(out);
  SpreadsheetApp.flush();
}

/** Statuses that mean "this recipient already got (or is getting) a certificate". */
function isDone(status) {
  return /^(sent|partial|sending)/i.test(status);
}

/** Builds a duplicate-detection key: same recipient(s) + same name. */
function makeKey(email, name) {
  const emails = extractEmails(email).map(a => a.toLowerCase()).sort().join(",");
  return emails + "|" + normalizeName(name);
}

/* ---------- Main ---------- */

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

    const hasWork = data.some(r => {
      const email = (r[0] || "").toString().trim();
      const status = (r[2] || "").toString().trim();
      return email && isPending(status);
    });
    if (!hasWork) return;

    // Register every row that was already sent, so new copies are caught as duplicates.
    const seen = {};
    for (let i = 0; i < data.length; i++) {
      const email = (data[i][0] || "").toString().trim();
      const name = (data[i][1] || "").toString().trim();
      const status = (data[i][2] || "").toString().trim();
      if (email && name && isDone(status)) {
        const key = makeKey(email, name);
        if (!seen[key]) seen[key] = i + 2;
      }
    }

    const draftToSend = findDraft(CONFIG.DRAFT_SUBJECT);
    if (!draftToSend) {
      Logger.log("Draft not found - nothing sent this run.");
      markNotSent(sheet, `draft "${CONFIG.DRAFT_SUBJECT}" not found`, 2);
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

    _pdfMap = null;

    for (let idx = 0; idx < data.length; idx++) {
      if (Date.now() - startTime > CONFIG.MAX_RUNTIME_MS) {
        Logger.log("Time limit near - stopping. Remaining rows will be sent on the next change/run.");
        markNotSent(sheet, "time limit reached, run again", idx + 2);
        break;
      }

      const rowNum = idx + 2;
      const email = (data[idx][0] || "").toString().trim();
      const name = (data[idx][1] || "").toString().trim();
      const status = (data[idx][2] || "").toString().trim();

      if (!email) continue;
      if (!isPending(status)) continue;

      if (!name) {
        setStatus(sheet, rowNum, "Failed: name is empty");
        continue;
      }

      const allParts = email.split(/[,;]/).map(a => a.trim()).filter(Boolean);
      const emailList = allParts.filter(a => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(a));
      const invalidList = allParts.filter(a => !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(a));

      if (emailList.length === 0) {
        setStatus(sheet, rowNum, "Failed: no valid email");
        continue;
      }

      // Duplicate check (skipped when the status is "Resend")
      const key = makeKey(email, name);
      if (!/^(resend|not sent \(resend\))/i.test(status)) {
        if (seen[key]) {
          setStatus(sheet, rowNum, `Duplicate of row ${seen[key]}`);
          continue;
        }
      }
      if (!seen[key]) seen[key] = rowNum;

      // Quota: every To, Cc and Bcc address counts.
      const perEmail = 1 + extractEmails(draftCc).length + extractEmails(draftBcc).length;
      if (MailApp.getRemainingDailyQuota() < emailList.length * perEmail) {
        Logger.log("Daily email quota exhausted - stopping.");
        markNotSent(sheet, "daily email quota exhausted", rowNum);
        break;
      }

      const customPdf = getCustomCertificate(name);
      if (!customPdf) {
        setStatus(sheet, rowNum, `Failed: PDF missing for '${name}.pdf'`);
        Logger.log(`Skipped ${name}: certificate not found in Drive folder.`);
        continue;
      }

      // Mark before sending so a timeout/crash can't cause a re-send.
      setStatus(sheet, rowNum, "Sending...");
      sendRowEmails(sheet, rowNum, emailList, invalidList, name, message, customPdf, extraAttachments, draftCc, draftBcc);
      Utilities.sleep(CONFIG.SEND_DELAY_MS);
    }
  } catch (err) {
    Logger.log("Unexpected error: " + err.message);
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
      if (sheet) markNotSent(sheet, "error: " + err.message, 2);
    } catch (e2) {
      Logger.log("Could not write Not sent status: " + e2.message);
    }
  } finally {
    lock.releaseLock();
  }
}

function setStatus(sheet, row, value) {
  sheet.getRange(row, 3).setValue(value);
  SpreadsheetApp.flush();
}

/* ---------- PDF lookup ---------- */

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
      if (_pdfMap[key]) {
        Logger.log(`Warning: multiple PDFs match "${key}". Using the first one.`);
      } else {
        _pdfMap[key] = f;
      }
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

function extractEmails(str) {
  if (!str) return [];
  const found = str.toString().match(/[A-Za-z0-9._%+\-']+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/g);
  return found || [];
}

function getRawHeader(raw, headerName) {
  const headerPart = raw.split(/\r?\n\r?\n/)[0];
  const unfolded = headerPart.replace(/\r?\n[ \t]+/g, " ");
  const re = new RegExp("^" + headerName + ":\\s*(.*)$", "im");
  const m = unfolded.match(re);
  return m ? m[1] : "";
}

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

  // Method 4: Gmail REST API using the script's own token
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

  cc = mergeEmails(cc, CONFIG.FALLBACK_CC);
  bcc = mergeEmails(bcc, CONFIG.FALLBACK_BCC);

  if (!cc && !bcc) {
    debug.push("WARNING: no Cc/Bcc found anywhere. Set FALLBACK_CC / FALLBACK_BCC in CONFIG.");
  }

  return { cc: cc, bcc: bcc, debug: debug };
}

/* ---------- Placeholders ---------- */

function escapeHtml(s) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
          .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

/** Function replacer so "$&" or "$1" in a name is never treated as a pattern. */
function fillPlaceholders(text, name, asHtml) {
  const value = asHtml ? escapeHtml(name) : name;
  return text.replace(/{{\s*name\s*}}/gi, () => value);
}

/* ---------- Inline images (logos, signatures) ---------- */

/** Maps each cid: reference in the HTML to an inline image blob, in order. */
function buildInlineImages(message, html) {
  const blobs = message.getAttachments({ includeInlineImages: true, includeAttachments: false });
  const cids = [];
  const re = /src=["']cid:([^"']+)["']/gi;
  let m;
  while ((m = re.exec(html)) !== null) {
    if (cids.indexOf(m[1]) === -1) cids.push(m[1]);
  }
  const map = {};
  for (let i = 0; i < Math.min(cids.length, blobs.length); i++) {
    map[cids[i]] = blobs[i];
  }
  return map;
}

/* ---------- Sending ---------- */

function sendRowEmails(sheet, row, emailList, invalidList, name, message, customPdf, extraAttachments, draftCc, draftBcc) {
  // Regular attachments only (inline images are handled separately)
  const draftAttachments = message.getAttachments({ includeInlineImages: false });
  const attachments = [customPdf].concat(draftAttachments).concat(extraAttachments || []);

  const rawHtml = message.getBody();
  const inlineImages = buildInlineImages(message, rawHtml);

  const subject = fillPlaceholders(message.getSubject(), name, false);
  const plainBody = fillPlaceholders(message.getPlainBody(), name, false);
  const htmlBody = fillPlaceholders(rawHtml, name, true);

  const succeeded = [];
  const failed = [];

  emailList.forEach(recipient => {
    try {
      const mailOptions = {
        htmlBody: htmlBody,
        attachments: attachments,
        name: CONFIG.SENDER_NAME,
      };
      if (Object.keys(inlineImages).length > 0) mailOptions.inlineImages = inlineImages;
      if (draftCc) mailOptions.cc = draftCc;
      if (draftBcc) mailOptions.bcc = draftBcc;

      GmailApp.sendEmail(recipient, subject, plainBody, mailOptions);
      succeeded.push(recipient);
      Logger.log(`Sent certificate to ${name} <${recipient}> | Cc="${draftCc}" Bcc="${draftBcc}"`);
    } catch (err) {
      failed.push(recipient + " (" + err.message + ")");
      Logger.log(`Failed for ${name} <${recipient}>: ` + err.message);
    }
  });

  let statusValue;
  if (failed.length === 0) {
    statusValue = "Sent";
  } else if (succeeded.length === 0) {
    statusValue = "Failed: " + failed.join("; ");
  } else {
    statusValue = "Partial - sent to " + succeeded.join(", ") + "; failed: " + failed.join("; ");
  }
  if (invalidList && invalidList.length > 0 && succeeded.length > 0) {
    statusValue += ` (skipped invalid: ${invalidList.join(", ")})`;
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
