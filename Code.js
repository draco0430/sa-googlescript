function saTicketRunCreateTicketsAndLogOnePerThread() {
  const scriptProps = PropertiesService.getScriptProperties();

  // ---- time window: (lastRun -> now) with optional overlap ----
  const currentTime = new Date();
  const currentTimeMs = currentTime.getTime();

  const runEveryMinutes = SA_TICKET_APP_CONFIG.RUN_EVERY_MINUTES || 5;
  const overlapMinutes = SA_TICKET_APP_CONFIG.OVERLAP_MINUTES || 0;

  const lastRunPropertyKey = SA_TICKET_APP_CONFIG.SCRIPT_PROPERTY_KEYS.LAST_RUN_MS;
  const doneThreadPrefix = SA_TICKET_APP_CONFIG.SCRIPT_PROPERTY_KEYS.DONE_THREAD_PREFIX;

  const lastRunMsRaw = scriptProps.getProperty(lastRunPropertyKey);
  const defaultWindowStartMs = currentTimeMs - (runEveryMinutes * 60 * 1000);

  // If first run ever, only process the last RUN_EVERY_MINUTES (not retroactive)
  const lastRunMs = lastRunMsRaw ? parseInt(lastRunMsRaw, 10) : defaultWindowStartMs;

  // Apply overlap by moving start earlier (optional)
  const windowStartMs = lastRunMs - (overlapMinutes * 60 * 1000);
  const windowEndMs = currentTimeMs;

  // Load email -> entity map from accounts sheet
  const emailToEntityMap = saTicketLoadEmailToEntityMap_();

  // Search inbox threads, not just unread
  const inboxThreads = GmailApp.search(
    "in:inbox",
    0,
    SA_TICKET_APP_CONFIG.MAX_THREADS_PER_RUN
  );

  if (!inboxThreads || inboxThreads.length === 0) {
    scriptProps.setProperty(lastRunPropertyKey, String(currentTimeMs));
    return;
  }

  const ticketLogSheet = saTicketGetOrCreateTicketLogSheet_();

  try {
    for (const gmailThread of inboxThreads) {
      const gmailThreadId = gmailThread.getId();
      const doneThreadKey = doneThreadPrefix + gmailThreadId;

      const newestMessage = saTicketGetNewestMessageInThread_(gmailThread);
      if (!newestMessage) continue;

      const messageDate = newestMessage.getDate();
      const messageMs = messageDate.getTime();

      const doneThreadValue = scriptProps.getProperty(doneThreadKey);

      // Existing processed thread:
      // if a newer message came in and thread is still Ticketed, add note and switch to Pending
      if (doneThreadValue) {
        const processedMs = Date.parse(doneThreadValue);

        if (
          !isNaN(processedMs) &&
          messageMs > processedMs &&
          saTicketThreadHasLabel_(gmailThread, "Ticketed")
        ) {
          const savedTicketInfo = saTicketGetSavedTicketInfoForThread_(gmailThreadId);

          if (savedTicketInfo) {
            const replyBody = saTicketGetReplyOnlyBody_(newestMessage);

            saTicketAddReplyNoteToTicket_(
              savedTicketInfo.ticketId,
              savedTicketInfo.ticketNumber,
              replyBody
            );

            saTicketRemoveLabelFromThread_(gmailThread, "Ticketed");
            saTicketApplyLabelToThread_(gmailThread, "Pending");

            // advance processed timestamp so the same latest message is not added again
            scriptProps.setProperty(doneThreadKey, new Date().toISOString());

            Logger.log(
              `NOTE ADDED thread=${gmailThreadId} ticketId=${savedTicketInfo.ticketId} ticketNumber=${savedTicketInfo.ticketNumber}`
            );
          } else {
            Logger.log(
              `SKIP NOTE thread=${gmailThreadId} reason=no saved ticket info found`
            );
          }
        }

        continue;
      }

      // Only process brand new threads/messages received within (lastRun -> now)
      if (!(messageMs > windowStartMs && messageMs <= windowEndMs)) {
        continue;
      }

      const gmailMessageId = newestMessage.getId();
      const senderEmail = saTicketExtractEmail_(newestMessage.getFrom() || "").toLowerCase();

      // Exclusion
      if (saTicketIsExcludedSender_(senderEmail)) {
        scriptProps.setProperty(doneThreadKey, new Date().toISOString());
        newestMessage.markRead();
        saTicketMarkThreadMessagesRead_(gmailThread);
        saTicketApplyLabelToThread_(gmailThread, "Ticketed");
        continue;
      }

      const originalSubject = (newestMessage.getSubject() || "").trim() || "(no subject)";
      const ticketBody = saTicketGetBestBody_(newestMessage);
      const ticketSubject = saTicketBuildUrgentSubject_(originalSubject, ticketBody);

      const entityIdToUse =
        emailToEntityMap[senderEmail] || SA_TICKET_APP_CONFIG.DEFAULT_ENTITY_ID_IF_NOT_FOUND;

      // No account match: notify Ian, but do not mark read and do not label Ticketed
      if (!entityIdToUse) {
        saTicketSendNoMatchEmail_(senderEmail, originalSubject, ticketBody);
        scriptProps.setProperty(doneThreadKey, new Date().toISOString());
        saTicketApplyLabelToThread_(gmailThread, "No Account");
        Logger.log(
          `NO MATCH thread=${gmailThreadId} sender=${senderEmail} subject=${originalSubject}`
        );
        continue;
      }

      // 1) Create ticket
      const createTicketResponse = saTicketCreateTicket_(ticketSubject, ticketBody, entityIdToUse);
      const createdTicketId =
        createTicketResponse && createTicketResponse.ID
          ? String(createTicketResponse.ID)
          : "";

      if (!createdTicketId) {
        throw new Error("Create ticket response missing ID");
      }

      // 2) Get ticket number
      const ticketOverlayResponse = saTicketGetTicketOverlay_(createdTicketId);
      const createdTicketNumber =
        ticketOverlayResponse && ticketOverlayResponse.TicketNumber != null
          ? String(ticketOverlayResponse.TicketNumber)
          : "";

      // 3) Log it
      saTicketAppendTicketLogRow_(ticketLogSheet, {
        timestamp: new Date(),
        gmailMessageId: gmailMessageId,
        threadId: gmailThreadId,
        fromEmail: senderEmail,
        subject: ticketSubject,
        ticketId: createdTicketId,
        ticketNumber: createdTicketNumber,
        entityIdUsed: entityIdToUse
      });

      // 4) Mark thread processed and save ticket mapping
      scriptProps.setProperty(doneThreadKey, new Date().toISOString());
      saTicketSaveThreadTicketInfo_(gmailThreadId, createdTicketId, createdTicketNumber);

      // 5) Mark read and label
      newestMessage.markRead();
      saTicketMarkThreadMessagesRead_(gmailThread);
      saTicketApplyLabelToThread_(gmailThread, "Ticketed");

      Logger.log(
        `OK thread=${gmailThreadId} msgDate=${messageDate.toISOString()} ticketId=${createdTicketId} ticketNumber=${createdTicketNumber} subject=${ticketSubject}`
      );
    }
  } finally {
    // Always advance last run time, even if an error occurs
    scriptProps.setProperty(lastRunPropertyKey, String(currentTimeMs));
  }

  // Optional cleanup:
  // saTicketCleanupProcessedThreadKeys_(60);
}

/* -------------------- Thread helpers -------------------- */

function saTicketGetNewestMessageInThread_(gmailThread) {
  const messages = gmailThread.getMessages();
  if (!messages || messages.length === 0) return null;
  return messages[messages.length - 1];
}

function saTicketMarkThreadMessagesRead_(gmailThread) {
  const messages = gmailThread.getMessages();
  for (const message of messages) {
    if (message.isUnread()) message.markRead();
  }
}

/* -------------------- Gmail label helpers -------------------- */

function saTicketApplyLabelToThread_(gmailThread, labelName) {
  if (!gmailThread || !labelName) return;

  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
  }

  gmailThread.addLabel(label);
}

function saTicketRemoveLabelFromThread_(gmailThread, labelName) {
  if (!gmailThread || !labelName) return;

  const label = GmailApp.getUserLabelByName(labelName);
  if (!label) return;

  gmailThread.removeLabel(label);
}

function saTicketThreadHasLabel_(gmailThread, labelName) {
  if (!gmailThread || !labelName) return false;

  const labels = gmailThread.getLabels();
  return labels.some(label => label.getName() === labelName);
}

/* -------------------- Sheet + logging -------------------- */

function saTicketLoadEmailToEntityMap_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const accountsSheet = spreadsheet.getSheetByName(SA_TICKET_APP_CONFIG.ACCOUNTS_SHEET_NAME);

  if (!accountsSheet) {
    throw new Error(`Sheet not found: ${SA_TICKET_APP_CONFIG.ACCOUNTS_SHEET_NAME}`);
  }

  const lastRow = accountsSheet.getLastRow();
  if (lastRow < 2) return {};

  const rows = accountsSheet.getRange(2, 1, lastRow - 1, 4).getValues();

  const emailToEntityMap = {};
  for (const row of rows) {
    const emailAddress = (row[2] || "").toString().trim().toLowerCase();
    const entityId = (row[3] || "").toString().trim();
    if (emailAddress) emailToEntityMap[emailAddress] = entityId;
  }

  return emailToEntityMap;
}

function saTicketGetOrCreateTicketLogSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let ticketLogSheet = spreadsheet.getSheetByName(SA_TICKET_APP_CONFIG.TICKET_LOG_SHEET_NAME);

  if (!ticketLogSheet) {
    ticketLogSheet = spreadsheet.insertSheet(SA_TICKET_APP_CONFIG.TICKET_LOG_SHEET_NAME);
  }

  const logHeaders = [
    "Timestamp",
    "Gmail Message ID",
    "Thread ID",
    "From email",
    "Subject",
    "Ticket ID",
    "Ticket number",
    "Entity ID used"
  ];

  if (ticketLogSheet.getLastRow() === 0) {
    ticketLogSheet.appendRow(logHeaders);
  } else {
    const existingHeaders = ticketLogSheet.getRange(1, 1, 1, logHeaders.length).getValues()[0];
    const existingHeadersJoined = existingHeaders.map(String).join("|");
    const expectedHeadersJoined = logHeaders.join("|");

    if (existingHeadersJoined !== expectedHeadersJoined) {
      ticketLogSheet.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]);
    }
  }

  return ticketLogSheet;
}

function saTicketAppendTicketLogRow_(ticketLogSheet, logRow) {
  ticketLogSheet.appendRow([
    logRow.timestamp,
    logRow.gmailMessageId,
    logRow.threadId,
    logRow.fromEmail,
    logRow.subject,
    logRow.ticketId,
    logRow.ticketNumber,
    logRow.entityIdUsed
  ]);
}

/* -------------------- SA API calls -------------------- */

function saTicketCreateTicket_(subject, body, entityId) {
  const requestPayload = {
    Ticket: {
      CategoryID: null,
      TicketStatus: SA_TICKET_APP_CONFIG.TICKET_STATUS,
      EntityID: entityId,
      EntityType: SA_TICKET_APP_CONFIG.ENTITY_TYPE,
      AssignmentID: SA_TICKET_APP_CONFIG.ASSIGNMENT_ID,
      DueDate: "",
      TicketDetail: {
        TicketEventType: SA_TICKET_APP_CONFIG.TICKET_EVENT_TYPE,
        Subject: subject,
        Body: body
      }
    }
  };

  return saTicketPostJson_(
    SA_TICKET_APP_CONFIG.CREATE_TICKET_ENDPOINT,
    requestPayload,
    "https://my.serviceautopilot.com/v3/CRM/TicketList"
  );
}

function saTicketGetTicketOverlay_(ticketId) {
  const requestPayload = { ticketId: ticketId };

  return saTicketPostJson_(
    SA_TICKET_APP_CONFIG.TICKET_OVERLAY_ENDPOINT,
    requestPayload,
    "https://my.serviceautopilot.com/v3/CRM/TicketList"
  );
}

function saTicketPostJson_(requestUrl, payloadObject, refererUrl) {
  if (
    !SA_TICKET_APP_CONFIG.COOKIE ||
    SA_TICKET_APP_CONFIG.COOKIE === "PASTE_COOKIE_HERE"
  ) {
    throw new Error("Config error: SA_TICKET_APP_CONFIG.COOKIE is not set.");
  }

  const requestHeaders = {
    "accept": "application/json, text/javascript, */*; q=0.01",
    "content-type": "application/json; charset=UTF-8",
    "origin": "https://my.serviceautopilot.com",
    "referer": refererUrl,
    "x-requested-with": "XMLHttpRequest",
    "cookie": SA_TICKET_APP_CONFIG.COOKIE
  };

  const response = UrlFetchApp.fetch(requestUrl, {
    method: "post",
    headers: requestHeaders,
    payload: JSON.stringify(payloadObject),
    muteHttpExceptions: true
  });

  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode < 200 || responseCode >= 300) {
    throw new Error(`HTTP ${responseCode}: ${responseText}`);
  }

  return JSON.parse(responseText);
}

/* -------------------- Gmail utils -------------------- */

function saTicketIsExcludedSender_(senderEmailLower) {
  const excludePatterns = SA_TICKET_APP_CONFIG.EXCLUDE_SENDER_PATTERNS || [];

  for (const pattern of excludePatterns) {
    const normalizedPattern = (pattern || "").toString().trim().toLowerCase();
    if (!normalizedPattern) continue;
    if (senderEmailLower.includes(normalizedPattern)) return true;
  }

  return false;
}

function saTicketExtractEmail_(fromRaw) {
  const emailMatch = fromRaw.match(/<([^>]+)>/);
  return (emailMatch ? emailMatch[1] : fromRaw).trim();
}

function saTicketGetBestBody_(gmailMessage) {
  let emailBody = (gmailMessage.getPlainBody() || "").trim();

  if (!emailBody) {
    emailBody = saTicketStripHtml_(gmailMessage.getBody() || "");
  }

  emailBody = emailBody.replace(/\r/g, "").trim();

  const maxBodyChars = SA_TICKET_APP_CONFIG.MAX_BODY_CHARS || 8000;
  if (emailBody.length > maxBodyChars) {
    emailBody = emailBody.slice(0, maxBodyChars) + "\n\n[TRUNCATED]";
  }

  return emailBody;
}

function saTicketStripHtml_(html) {
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<\/?[^>]+(>|$)/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/* -------------------- Optional cleanup -------------------- */

function saTicketCleanupProcessedThreadKeys_(keepDays) {
  const scriptProps = PropertiesService.getScriptProperties();
  const allProperties = scriptProps.getProperties();
  const doneThreadPrefix = SA_TICKET_APP_CONFIG.SCRIPT_PROPERTY_KEYS.DONE_THREAD_PREFIX;
  const cutoffMs = Date.now() - keepDays * 24 * 60 * 60 * 1000;

  for (const [propertyKey, propertyValue] of Object.entries(allProperties)) {
    if (!propertyKey.startsWith(doneThreadPrefix)) continue;

    const processedTimestamp = Date.parse(propertyValue);
    if (!isNaN(processedTimestamp) && processedTimestamp < cutoffMs) {
      scriptProps.deleteProperty(propertyKey);
    }
  }
}

/* -------------------- Trigger -------------------- */

function saTicketCreateTriggerEvery5Minutes() {
  ScriptApp.newTrigger("saTicketRunCreateTicketsAndLogOnePerThread")
    .timeBased()
    .everyMinutes(5)
    .create();
}

function saTicketInitLastRunNow() {
  PropertiesService.getScriptProperties().setProperty(
    SA_TICKET_APP_CONFIG.SCRIPT_PROPERTY_KEYS.LAST_RUN_MS,
    String(new Date().getTime())
  );
}

function saTicketSendNoMatchEmail_(senderEmail, originalSubject, emailBody) {
  const to = "ianjaylog@cloudstaff.com";
  const link = saAccountGenerateMagicLink_({
    email: senderEmail,
    ticketSubject: originalSubject,
    ticketBody: emailBody
  });

  const safeSenderEmail = senderEmail || "";
  const safeSubject = originalSubject || "(no subject)";
  const safeBody = (emailBody || "").trim();

  const plainBody =
    `Email: ${safeSenderEmail}\n` +
    `Subject: ${safeSubject}\n\n` +
    `Body:\n${safeBody}\n\n` +
    `Create account here (valid for 2 days):\n${link}\n`;

  const htmlBody =
    `<p><strong>Email:</strong> ${saTicketEscapeHtml_(safeSenderEmail)}</p>` +
    `<p><strong>Subject:</strong> ${saTicketEscapeHtml_(safeSubject)}</p>` +
    `<p><strong>Body:</strong><br>${saTicketEscapeHtml_(safeBody).replace(/\n/g, "<br>")}</p>` +
    `<p><strong>Create account here (valid for 2 days):</strong><br>` +
    `<a href="${link}" target="_blank">${link}</a></p>`;

  Logger.log("NO MATCH LINK=" + link);
  Logger.log("NO MATCH EMAIL BODY=" + plainBody);

  GmailApp.sendEmail(to, `NO MATCH: ${safeSubject}`, plainBody, {
    htmlBody: htmlBody
  });
}

function saTicketEscapeHtml_(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function testSendNoMatchEmail() {
  saTicketSendNoMatchEmail_(
    "ianjaylog@cloudstaff.com",
    "ppootest",
    "tesatta"
  );
}

function testMagicLink() {
  const link = saAccountGenerateMagicLink_({
    email: "ianjaylog@cloudstaff.com",
    ticketSubject: "Test subject",
    ticketBody: "Test body"
  });
  Logger.log("MAGIC LINK: " + link);
}


//test joke