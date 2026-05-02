function saAccountCreateInSA_(formData) {
  const payload = {
    Input: {
      IsLead: formData.accountType === "Lead",
      FirstName: formData.firstName || "",
      LastName: formData.lastName || "",
      Email: formData.email || "",
      PhonePrimary: formData.phonePrimary || "",
      PrimaryContact: Number(formData.primaryContact || 3),
      ServiceAddress: formData.serviceAddress || "",
      City: formData.city || "",
      State: formData.state || "",
      Zip: formData.zip || "",
      MapCode: formData.mapCode || "",
      IsTaxable: !!formData.isTaxable,
      AccountNumber: formData.accountNumber || "",
      ClientName: formData.displayName || "",
      TaxReference: formData.taxReference || "",
      NameOnInvoice: formData.nameOnInvoice || "",
      PropertyName: formData.propertyName || ""
    }
  };

  const response = saTicketPostJson_(
    "https://my.serviceautopilot.com/v3/Webservices/CRM/ClientsWS.asmx/SaveAccount",
    payload,
    "https://my.serviceautopilot.com/v3/CRM/Accounts"
  );

  // Adjust this after you inspect the real response shape from SA
  const entityId =
    (response && response.d && response.d.ID) ||
    (response && response.ID) ||
    (response && response.EntityID) ||
    "";

  return {
    entityId: String(entityId || ""),
    raw: response
  };
}

function saAccountAppendToAccountsSheet_(formData, created) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SA_TICKET_APP_CONFIG.ACCOUNTS_SHEET_NAME);

  if (!sheet) {
    throw new Error(`Sheet not found: ${SA_TICKET_APP_CONFIG.ACCOUNTS_SHEET_NAME}`);
  }

  // Assumes your current accounts sheet uses:
  // Col A = First Name
  // Col B = Last Name
  // Col C = Email
  // Col D = Entity ID
  sheet.appendRow([
    formData.firstName || "",
    formData.lastName || "",
    formData.email || "",
    created.entityId || ""
  ]);
}