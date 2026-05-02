function doGet(e) {
  const token = (e && e.parameter && e.parameter.token) ? e.parameter.token : "";
  const email = (e && e.parameter && e.parameter.email) ? e.parameter.email : "";

  const tokenCheck = saAccountValidateMagicLinkToken_(token);

  if (!tokenCheck.ok) {
    return HtmlService.createHtmlOutput(`
      <html>
        <body style="font-family: Arial, sans-serif; padding: 24px;">
          <h2>Link expired or invalid</h2>
          <p>This account creation link is no longer valid.</p>
        </body>
      </html>
    `).setTitle("Create Account");
  }

  const template = HtmlService.createTemplateFromFile("account_form");
  template.prefillEmail = email || tokenCheck.email || "";
  template.prefillTicketSubject = tokenCheck.ticketSubject || "";
  template.prefillTicketBody = tokenCheck.ticketBody || "";
  template.token = token;

  return template.evaluate()
    .setTitle("Create Account")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function saAccountGenerateMagicLink_(input) {
  const payload = (typeof input === "object" && input !== null)
    ? input
    : { email: input || "" };

  const token = Utilities.getUuid();
  const expiresAtMs = Date.now() + (2 * 24 * 60 * 60 * 1000); // 2 days

  const props = PropertiesService.getScriptProperties();
  props.setProperty(
    `SA_ACCOUNT_MAGIC_${token}`,
    JSON.stringify({
      email: payload.email || "",
      ticketSubject: payload.ticketSubject || "",
      ticketBody: payload.ticketBody || "",
      expiresAtMs: expiresAtMs
    })
  );

  const webAppUrl = ScriptApp.getService().getUrl();
  return `${webAppUrl}?token=${encodeURIComponent(token)}&email=${encodeURIComponent(payload.email || "")}`;
}

function saAccountValidateMagicLinkToken_(token) {
  if (!token) {
    return { ok: false, reason: "missing token" };
  }

  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(`SA_ACCOUNT_MAGIC_${token}`);
  if (!raw) {
    return { ok: false, reason: "token not found" };
  }

  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch (err) {
    return { ok: false, reason: "invalid token payload" };
  }

  if (!parsed.expiresAtMs || Date.now() > parsed.expiresAtMs) {
    return { ok: false, reason: "token expired" };
  }

  return {
    ok: true,
    email: parsed.email || "",
    ticketSubject: parsed.ticketSubject || "",
    ticketBody: parsed.ticketBody || ""
  };
}

function saAccountConsumeMagicLinkToken_(token) {
  if (!token) return;
  PropertiesService.getScriptProperties().deleteProperty(`SA_ACCOUNT_MAGIC_${token}`);
}

function saAccountHandleCreateFromForm_(formData) {
  const tokenCheck = saAccountValidateMagicLinkToken_(formData.token);
  if (!tokenCheck.ok) {
    throw new Error("This link is expired or invalid.");
  }

  const created = saAccountCreateInSA_(formData);
  saAccountAppendToAccountsSheet_(formData, created);

  saAccountConsumeMagicLinkToken_(formData.token);

  return {
    ok: true,
    message: "Account created successfully.",
    entityId: created.entityId || "",
    raw: created.raw || null
  };
}