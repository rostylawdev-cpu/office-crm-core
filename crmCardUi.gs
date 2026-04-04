/** crmCardUi.gs */

function crm_openMatterCardFromActiveRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  const row = sh.getActiveRange().getRow();
  const matterId = String(sh.getRange(row, 1).getValue() || "").trim();

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
    limitActivities: 50,
    limitDocuments: 50,
  });

  if (!card || !card.ok) {
    SpreadsheetApp.getUi().alert(`Дело не найдено: ${matterId}`);
    return;
  }

  const t = HtmlService.createTemplateFromFile("crm_matter_card");
  t.card = card;
  t.matterId = matterId;

  const html = t
    .evaluate()
    .setTitle("CRM Matter Card")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  SpreadsheetApp.getUi().showSidebar(html);
}

/** ===== Matter actions from UI ===== */

function crm_uiUpdateMatterStage(matterId, stage) {
  return crm_updateMatterStage(matterId, stage, "Updated from Matter Card UI");
}

function crm_uiUpdateMatterEventDate(matterId, eventDate) {
  if (!matterId) throw new Error("crm_uiUpdateMatterEventDate: matterId is required");
  crm_setMatterField(matterId, "EVENT_DATE", String(eventDate || "").trim());
  return { ok: true };
}

function crm_uiAddTaskToMatter(payload) {
  if (!payload || !payload.matterId) throw new Error("crm_uiAddTaskToMatter: missing matterId");

  const matter = crm_getMatterById(payload.matterId);
  if (!matter) throw new Error("crm_uiAddTaskToMatter: matter not found");

  return crm_createTask({
    clientId: matter.CLIENT_ID,
    matterId: payload.matterId,
    type: payload.type || "GENERAL",
    title: payload.title || "",
    dueDate: payload.dueDate || "",
    status: "OPEN",
    priority: payload.priority || "MEDIUM",
    assignee: payload.assignee || getActiveUserEmail_() || "unknown",
    generatedBy: "MATTER_CARD_UI",
    notes: payload.notes || "",
  });
}

function crm_uiMarkTaskDone(taskId, note) {
  return crm_markTaskDone(taskId, note || "Done from Matter Card UI");
}

function crm_uiAddDocumentToMatter(payload) {
  if (!payload || !payload.matterId) throw new Error("crm_uiAddDocumentToMatter: missing matterId");

  const matter = crm_getMatterById(payload.matterId);
  if (!matter) throw new Error("crm_uiAddDocumentToMatter: matter not found");

  return crm_addDocument({
    clientId: matter.CLIENT_ID,
    matterId: payload.matterId,
    type: payload.type || "GENERAL",
    status: payload.status || "DRAFT",
    title: payload.title || "",
    docUrl: payload.docUrl || "",
    pdfUrl: payload.pdfUrl || "",
    fileId: payload.fileId || "",
    createdBy: getActiveUserEmail_() || "unknown",
    notes: payload.notes || "",
  });
}

function crm_uiGenerateAgreementPoa(matterId) {
  if (!matterId) throw new Error("crm_uiGenerateAgreementPoa: matterId is required");
  return crm_generateAgreementAndPoa(matterId);
}

function crm_uiCreateSignLinksForMatter(matterId) {
  if (!matterId) throw new Error("crm_uiCreateSignLinksForMatter: matterId is required");
  return crm_createSignLinksForMatter(matterId);
}

function crm_uiUploadDocumentToMatter(payload) {
  return crm_uploadFileToDriveAndRegister(payload);
}