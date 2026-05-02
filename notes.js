function saTicketGetThreadTicketIdKey_(threadId) {
  return `SA_TICKET_THREAD_TICKET_ID_${threadId}`;
}

function saTicketGetThreadTicketNumberKey_(threadId) {
  return `SA_TICKET_THREAD_TICKET_NUMBER_${threadId}`;
}

function saTicketSaveThreadTicketInfo_(threadId, ticketId, ticketNumber) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(saTicketGetThreadTicketIdKey_(threadId), String(ticketId));
  props.setProperty(saTicketGetThreadTicketNumberKey_(threadId), String(ticketNumber));
}

function saTicketGetSavedTicketInfoForThread_(threadId) {
  const props = PropertiesService.getScriptProperties();

  const ticketId = props.getProperty(saTicketGetThreadTicketIdKey_(threadId));
  const ticketNumber = props.getProperty(saTicketGetThreadTicketNumberKey_(threadId));

  if (!ticketId || !ticketNumber) return null;

  return {
    ticketId: ticketId,
    ticketNumber: Number(ticketNumber)
  };
}

function saTicketAddReplyNoteToTicket_(ticketId, ticketNumber, replyBody) {
  const payload = {
    Ticket: {
      TicketID: ticketId,
      TicketNumber: ticketNumber,
      EntityID: "00000000-0000-0000-0000-000000000000",
      EntityType: "",
      TicketStatus: SA_TICKET_APP_CONFIG.TICKET_STATUS,
      TicketDetail: {
        TicketDetailID: "00000000-0000-0000-0000-000000000000",
        TicketID: ticketId,
        TicketEventType: SA_TICKET_APP_CONFIG.TICKET_EVENT_TYPE,
        CreatedByID: SA_TICKET_APP_CONFIG.ASSIGNMENT_ID,
        CreatedByType: 2,
        Subject: "Customer Replied:",
        Body: replyBody,
        NumberOfAttachments: 0,
        Attachments: []
      }
    }
  };

  return saTicketPostJson_(
    SA_TICKET_APP_CONFIG.CREATE_TICKET_ENDPOINT,
    payload,
    "https://my.serviceautopilot.com/v3/CRM/TicketList"
  );
}

function saTicketGetReplyOnlyBody_(gmailMessage) {
  let body = (gmailMessage.getPlainBody() || "").replace(/\r/g, "").trim();

  if (!body) {
    body = saTicketStripHtml_(gmailMessage.getBody() || "");
  }

  const cutPatterns = [
    /^From:\s.*$/mi,
    /^Sent:\s.*$/mi,
    /^To:\s.*$/mi,
    /^Subject:\s.*$/mi,
    /^On .* wrote:$/mi,
    /^---+Original Message---+$/mi,
    /^_{5,}.*$/mi
  ];

  let cutIndex = body.length;

  for (const pattern of cutPatterns) {
    const match = body.match(pattern);
    if (match && match.index < cutIndex) {
      cutIndex = match.index;
    }
  }

  body = body.slice(0, cutIndex).trim();

  // Remove common mobile signatures / trailing separators if needed
  body = body
    .replace(/^Sent from my iPhone.*$/gmi, "")
    .replace(/^Sent from my Android.*$/gmi, "")
    .trim();

  return body;
}