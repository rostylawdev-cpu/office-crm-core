/** CrmSetup.js */

function crm_setPropsOnce() {
  PropertiesService.getScriptProperties().setProperty(
    "CRM_SPREADSHEET_ID",
    "1Py-Dht-1xBQ4vZ6zfiEQMmBZt327gIIKXrLholocUBw"
  );
  return "Script properties set";
}


function crm_setupCore() {
  const c = cfg_();

  const ssId = getProp_(c.CRM_SPREADSHEET_ID_PROP);
  if (!ssId) {
    throw new Error("Missing Script Property: CRM_SPREADSHEET_ID");
  }

  const ss = SpreadsheetApp.openById(ssId);

  const sheets = Object.values(c.SHEETS);
  sheets.forEach(name => ensureSheet_(ss, name));

  ensureHeaders_(ss.getSheetByName(c.SHEETS.CLIENTS), c.HEADERS.CLIENTS);
  ensureHeaders_(ss.getSheetByName(c.SHEETS.MATTERS), c.HEADERS.MATTERS);
  ensureHeaders_(ss.getSheetByName(c.SHEETS.DOCUMENTS), c.HEADERS.DOCUMENTS);
  ensureHeaders_(ss.getSheetByName(c.SHEETS.ACTIVITIES), c.HEADERS.ACTIVITIES);
  ensureHeaders_(ss.getSheetByName(c.SHEETS.TASKS), c.HEADERS.TASKS);

  logInfo_("SETUP", "CRM core scaffold created/verified", { spreadsheetId: ss.getId() });

  return "OK";
}

function test_addClient() {
  const res = crm_addClient({
    fullName: "Darwin Saindrom",
    phone: "050-1234567",
    email: "darwin@example.com",
    source: "TEST",
    locale: "ru-IL",
  });
  console.log(res);
}