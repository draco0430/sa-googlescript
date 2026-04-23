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

  // Find unread threads (broad), then filter strictly by msg.getDate() within window
  const unreadThreads = GmailApp.search(
    "in:inbox is:unread",
    0,
    SA_TICKET_APP_CONFIG.MAX_THREADS_PER_RUN
  );

  if (!unreadThreads || unreadThreads.length === 0) {
    scriptProps.setProperty(lastRunPropertyKey, String(currentTimeMs));
    return;
  }

  const ticketLogSheet = saTicketGetOrCreateTicketLogSheet_();

  try {
    for (const gmailThread of unreadThreads) {
      const gmailThreadId = gmailThread.getId();
      const doneThreadKey = doneThreadPrefix + gmailThreadId;

      // One ticket per thread
      if (scriptProps.getProperty(doneThreadKey)) {
        // Optional: mark replies read so it doesn't keep popping up
        saTicketMarkThreadMessagesRead_(gmailThread);
        continue;
      }

      const newestUnreadMessage = saTicketGetNewestUnreadMessageInThread_(gmailThread);
      if (!newestUnreadMessage) continue;

      // STRICT time window filter based on the received date/time
      const messageDate = newestUnreadMessage.getDate();
      const messageMs = messageDate.getTime();

      // Only process emails received within (lastRun -> now)
      if (!(messageMs > windowStartMs && messageMs <= windowEndMs)) {
        // Do NOT mark read and do NOT mark processed
        continue;
      }

      const gmailMessageId = newestUnreadMessage.getId();
      const senderEmail = saTicketExtractEmail_(newestUnreadMessage.getFrom() || "").toLowerCase();

      // Exclusion
      if (saTicketIsExcludedSender_(senderEmail)) {
        scriptProps.setProperty(doneThreadKey, new Date().toISOString());
        newestUnreadMessage.markRead();
        continue;
      }

      const ticketSubject = (newestUnreadMessage.getSubject() || "").trim() || "(no subject)";
      const ticketBody = saTicketGetBestBody_(newestUnreadMessage);

      const entityIdToUse =
        emailToEntityMap[senderEmail] || SA_TICKET_APP_CONFIG.DEFAULT_ENTITY_ID_IF_NOT_FOUND;

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

      // 4) Mark thread processed
      scriptProps.setProperty(doneThreadKey, new Date().toISOString());

      // Mark read
      newestUnreadMessage.markRead();
      saTicketMarkThreadMessagesRead_(gmailThread);

      Logger.log(
        `OK thread=${gmailThreadId} msgDate=${messageDate.toISOString()} ticketId=${createdTicketId} ticketNumber=${createdTicketNumber}`
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

function saTicketGetNewestUnreadMessageInThread_(gmailThread) {
  const messages = gmailThread.getMessages();
  for (let i = messages.length - 1; i >= 0; i--) {
    if (messages[i].isUnread()) return messages[i];
  }
  return null;
}

function saTicketMarkThreadMessagesRead_(gmailThread) {
  const messages = gmailThread.getMessages();
  for (const message of messages) {
    if (message.isUnread()) message.markRead();
  }
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