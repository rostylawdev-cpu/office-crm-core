/** crmClientsUi.gs */

function crm_openClientCardById(clientId) {
  if (!clientId) {
    SpreadsheetApp.getUi().alert("crm_openClientCardById: missing clientId");
    return;
  }
  crm_openClientCard_(clientId);
}

function crm_openClientCardFromActiveRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  const row = sh.getActiveRange().getRow();

  const clientId = getRowValueByHeader_(sh, row, "CLIENT_ID");

  if (!clientId) {
    SpreadsheetApp.getUi().alert("Не удалось найти CLIENT_ID в активной строке.");
    return;
  }

  crm_openClientCard_(String(clientId).trim());
}

function crm_openClientCard_(clientId) {
  const card = crm_getClientCard(clientId, {
    tasksStatus: "OPEN",
    limitTasks: 50,
    limitActivities: 50,
    limitDocuments: 50,
  });

  if (!card || !card.ok) {
    SpreadsheetApp.getUi().alert(`Клиент не найден: ${clientId}`);
    return;
  }

  const t = HtmlService.createTemplateFromFile("crm_client_card");
  t.card = card;
  t.clientId = clientId;

  const html = t
    .evaluate()
    .setTitle("CRM Client Card")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  SpreadsheetApp.getUi().showSidebar(html);
}

function crm_renderClientCardHtml(clientId) {
  const card = crm_getClientCard(clientId, {
    tasksStatus: "OPEN",
    limitTasks: 20,
    limitActivities: 20,
    limitDocuments: 20,
  });

  if (!card || !card.ok) {
    return HtmlService.createHtmlOutput(
      `<div style="font-family:Arial;padding:12px">
         <b>Ошибка:</b> ${escapeHtml_(card?.error || "UNKNOWN")}<br/>
         clientId: ${escapeHtml_(clientId || "")}
       </div>`
    ).setTitle("CRM Client Card");
  }

  const tpl = HtmlService.createTemplateFromFile("crm_client_card");
  tpl.card = card;
  tpl.clientId = clientId;

  return tpl
    .evaluate()
    .setTitle("CRM Client Card")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}