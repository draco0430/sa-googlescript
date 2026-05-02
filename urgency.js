function saTicketIsUrgent_(subject, body) {
  const textToScan = `${subject || ""}\n${body || ""}`.toLowerCase();

  const dangerKeywords = [
    "burn",
    "dead",
    "brown stripes",
    "tracks",
    "dead patches",
    "hit",
    "stuck",
    "broke",
    "killed",
    "blanket spray",
    "overspray",
    "sod",
    "ruined",
    "review",
    "1-star",
    "google review",
    "social media",
    "lawyer",
    "attorney",
    "pita",
    "unauthorized",
    "scam",
    "refund",
    "cancel",
    "stop service",
    "no response",
    "left messages",
    "waiting",
    "called back",
    "manager",
    "owner",
    "frustrated",
    "fed up",
    "irate",
    "help",
    "dog",
    "child",
    "chemicals",
    "sprayed",
    "technician",
    "specialist",
    "cursing",
    "rude",
    "yelled"
  ];

  return dangerKeywords.some(keyword => textToScan.includes(keyword));
}

function saTicketBuildUrgentSubject_(subject, body) {
  const cleanSubject = (subject || "").trim() || "(no subject)";

  if (!saTicketIsUrgent_(cleanSubject, body)) {
    return cleanSubject;
  }

  if (/\[URGENT\]\s*$/i.test(cleanSubject)) {
    return cleanSubject;
  }

  return `${cleanSubject} [URGENT]`;
}