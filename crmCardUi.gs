function crm_openMatterCardFromActiveRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  const row = sh.getActiveRange().getRow();
  const matterId = String(sh.getRange(row, 1).getValue() || "").trim(); // MATTER_ID в A

  if (!matterId) {
    SpreadsheetApp.getUi().alert("Выбери строку с MATTER_ID в колонке A.");
    return;
  }

  crm_openMatterCard_(matterId);
}

function crm_openMatterCardById(matterId) {
  crm_openMatterCard_(matterId);
}

function crm_openMatterCard_(matterId) {
  const card = crm_getMatterCard(matterId, {
    tasksStatus: "OPEN",
    limitTasks: 50,
    limitActivities: 50
  });

  const t = HtmlService.createTemplateFromFile("crm_matter_card");
  t.card = card;
  t.matterId = matterId;

  const html = t.evaluate()
    .setTitle("CRM Matter Card")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  SpreadsheetApp.getUi().showSidebar(html);
}