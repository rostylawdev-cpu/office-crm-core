/** crmTests.gs */

/** =========================
 *  CLIENTS
 *  ========================= */

function test_addClient() {
  const res = crm_addClient({
    fullName: "Darwin Saindrom",
    phone: "050-1234567",
    email: "darwin@example.com",
    source: "TEST",
    locale: "ru-IL",
  });

  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

function test_upsertClient() {
  const res = crm_upsertClient({
    phone: "050-1234567",
    email: "darwin+new@example.com",
    fullName: "Darwin Saindrom",
    source: "TEST_UPSERT",
  });

  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

function test_findClient() {
  const found = crm_findClientByPhone("050-1234567");
  Logger.log(JSON.stringify(found, null, 2));
  return found ? found.CLIENT_ID : "NOT_FOUND";
}

/** =========================
 *  MATTERS
 *  ========================= */

function test_createMatter() {
  const clientId = "CL_20260304_204602_6MLHYU"; // подставь свой реальный

  const res = crm_createMatter({
    clientId,
    category: "WORK_ACCIDENT",
    title: "Teunat Avoda – test matter",
    authority: "Bituach Leumi",
  });

  Logger.log(JSON.stringify(res, null, 2));

  const list = crm_findMattersByClientId(clientId);
  Logger.log(JSON.stringify(list, null, 2));

  return res;
}

function test_createMatterWithTasks() {
  const clientId = "CL_20260304_204602_6MLHYU"; // подставь свой реальный

  const res = crm_createMatterWithTasks({
    clientId,
    category: "WORK_ACCIDENT",
    title: "Teunat Avoda – workflow test",
    authority: "Bituach Leumi",
    taskTemplateKey: "WORK_ACCIDENT",
  });

  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

/** =========================
 *  TASKS
 *  ========================= */

function test_createTask() {
  const clientId = "CL_20260304_204602_6MLHYU"; // подставь свой
  const matterId = "MAT_20260305_803005"; // подставь свой

  const res = crm_createTask({
    clientId,
    matterId,
    type: "FOLLOW_UP",
    title: "Позвонить клиенту и запросить טופס 250",
    dueDate: nowIso_(),
    priority: "HIGH",
    notes: "Тестовая задача",
  });

  Logger.log(JSON.stringify(res, null, 2));

  const open = crm_findOpenTasksByClientId(clientId);
  Logger.log(JSON.stringify(open, null, 2));

  return res;
}

function test_markDone() {
  const tasks = crm_listTasks({ status: "OPEN", limit: 1 });
  if (!tasks.length) throw new Error("No OPEN tasks found");

  const task = tasks[0];
  const res = crm_markTaskDone(task.TASK_ID, "Закрыто автотестом");

  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

/** =========================
 *  DASHBOARD / CARDS
 *  ========================= */

function test_dashboard() {
  const clientId = "CL_20260304_204602_6MLHYU"; // подставь свой реальный
  const dash = crm_getClientDashboard(clientId);

  Logger.log(JSON.stringify(dash, null, 2));
  return dash;
}

function test_getMatterCard() {
  const matters = crm_findMattersByClientId("CL_20260304_204602_6MLHYU");
  if (!matters.length) throw new Error("No matters found for test client");

  const matterId = matters[0].MATTER_ID;

  const card = crm_getMatterCard(matterId, {
    tasksStatus: "OPEN",
    limitTasks: 50,
    limitActivities: 20,
  });

  Logger.log(JSON.stringify(card, null, 2));
  return card;
}

/** =========================
 *  LEADS
 *  ========================= */

function test_addLead() {
  const res = crm_addLead({
    fullName: "Test Lead One",
    phone: "050-7777777",
    email: "lead@example.com",
    source: "FACEBOOK",
    campaign: "WORK_ACCIDENT_MARCH",
    caseType: "WORK_ACCIDENT",
    description: "Упал на работе, нужна консультация",
    status: "NEW",
    notes: "Тестовый лид",
  });

  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

function test_convertLeadToClientAndMatter() {
  const leads = crm_listLeads({ limit: 1 });
  if (!leads.length) throw new Error("No leads found");

  const lead = leads[0];

  const res = crm_convertLeadToClientAndMatter(lead.LEAD_ID, {
    category: "WORK_ACCIDENT",
    title: "Matter created from lead",
    authority: "Bituach Leumi",
    taskTemplateKey: "WORK_ACCIDENT",
  });

  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

function test_getLeadCard() {
  const leads = crm_listLeads({ limit: 1 });
  if (!leads.length) throw new Error("No leads found");

  const leadId = leads[0].LEAD_ID;

  const card = crm_getLeadCard(leadId, {
    limitActivities: 20,
  });

  Logger.log(JSON.stringify(card, null, 2));
  return card;
}

function test_convertLeadToClient() {
  const leadRes = crm_addLead({
    fullName: "Lead Convert Test",
    phone: "050-8888888",
    email: "lead.convert@example.com",
    source: "TEST",
    campaign: "TEST_CONVERT",
    caseType: "GENERAL",
    description: "Test convert lead to client",
    status: "NEW",
    notes: "autotest",
  });

  const leadId = leadRes.leadId;

  const res = crm_convertLeadToClient(leadId);

  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

function debug_showLeadCardJson() {
  const leads = crm_listLeads({ limit: 1 });
  if (!leads.length) throw new Error("No leads found");

  const leadId = leads[0].LEAD_ID;
  const card = crm_getLeadCard(leadId, { limitActivities: 20 });

  SpreadsheetApp.getUi().alert(JSON.stringify({
    ok: card.ok,
    leadId: leadId,
    hasLead: !!card.lead,
    hasClient: !!card.client,
    mattersCount: (card.matters || []).length,
    activitiesCount: (card.activities || []).length
  }, null, 2));

  return card;
}

function test_getClientCard() {
  const leads = crm_listLeads({ limit: 10 });
  if (!leads.length) throw new Error("No leads found");

  let clientId = "";
  for (var i = 0; i < leads.length; i++) {
    if (leads[i].CLIENT_ID) {
      clientId = leads[i].CLIENT_ID;
      break;
    }
  }

  if (!clientId) {
    const found = crm_findClientByPhone("050-1234567");
    if (found && found.CLIENT_ID) {
      clientId = found.CLIENT_ID;
    }
  }

  if (!clientId) throw new Error("No clientId found for test_getClientCard");

  const card = crm_getClientCard(clientId, {
    tasksStatus: "OPEN",
    limitTasks: 20,
    limitActivities: 20,
  });

  Logger.log(JSON.stringify(card, null, 2));
  return card;
}

function test_searchAll() {
  const res = crm_searchAll("050-7777777", { limit: 10 });
  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

function test_addDocument() {
  const found = crm_findClientByPhone("050-7777777") || crm_findClientByPhone("050-1234567");
  if (!found || !found.CLIENT_ID) throw new Error("No test client found");

  const matters = crm_findMattersByClientId(found.CLIENT_ID) || [];
  const matterId = matters.length ? matters[0].MATTER_ID : "";

  const res = crm_addDocument({
    clientId: found.CLIENT_ID,
    matterId: matterId,
    type: "AGREEMENT",
    status: "READY",
    title: "Test Agreement Document",
    docUrl: "https://docs.google.com/document/d/test-doc-id/edit",
    pdfUrl: "https://drive.google.com/file/d/test-pdf-id/view",
    fileId: "test-file-id",
    notes: "Autotest document",
  });

  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

function test_listDocumentsByClient() {
  const found = crm_findClientByPhone("050-7777777") || crm_findClientByPhone("050-1234567");
  if (!found || !found.CLIENT_ID) throw new Error("No test client found");

  const docs = crm_listDocumentsByClientId(found.CLIENT_ID, { limit: 20 });
  Logger.log(JSON.stringify(docs, null, 2));
  return docs;
}

function test_listDocumentsByMatter() {
  const matters = crm_findMattersByClientId("CL_20260315_182218_GN32F7");
  if (!matters.length) throw new Error("No matters found for test client");

  const docs = crm_listDocumentsByMatterId(matters[0].MATTER_ID, { limit: 20 });
  Logger.log(JSON.stringify(docs, null, 2));
  return docs;
}

function test_updateMatterStage() {
  const found = crm_findClientByPhone("050-7777777") || crm_findClientByPhone("050-1234567");
  if (!found || !found.CLIENT_ID) throw new Error("No test client found");

  const matters = crm_findMattersByClientId(found.CLIENT_ID) || [];
  if (!matters.length) throw new Error("No matters found");

  const matterId = matters[0].MATTER_ID;
  const res = crm_updateMatterStage(matterId, "DOCS", "Autotest stage update");

  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

function test_addTaskToMatterManual() {
  const found = crm_findClientByPhone("050-7777777") || crm_findClientByPhone("050-1234567");
  if (!found || !found.CLIENT_ID) throw new Error("No test client found");

  const matters = crm_findMattersByClientId(found.CLIENT_ID) || [];
  if (!matters.length) throw new Error("No matters found");

  const matterId = matters[0].MATTER_ID;

  const res = crm_createTask({
    clientId: found.CLIENT_ID,
    matterId: matterId,
    type: "FOLLOW_UP",
    title: "Manual task from autotest",
    dueDate: nowIso_(),
    status: "OPEN",
    priority: "MEDIUM",
    assignee: getActiveUserEmail_() || "unknown",
    generatedBy: "AUTOTEST",
    notes: "Autotest manual task",
  });

  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

function test_addDocumentToMatterManual() {
  const found = crm_findClientByPhone("050-7777777") || crm_findClientByPhone("050-1234567");
  if (!found || !found.CLIENT_ID) throw new Error("No test client found");

  const matters = crm_findMattersByClientId(found.CLIENT_ID) || [];
  if (!matters.length) throw new Error("No matters found");

  const matterId = matters[0].MATTER_ID;

  const res = crm_addDocument({
    clientId: found.CLIENT_ID,
    matterId: matterId,
    type: "MEDICAL",
    status: "READY",
    title: "Manual document from autotest",
    docUrl: "",
    pdfUrl: "",
    fileId: "",
    notes: "Autotest matter document",
  });

  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

function test_documentsAfterUpload() {
  const found = crm_findClientByPhone("050-7777777") || crm_findClientByPhone("050-1234567");
  if (!found || !found.CLIENT_ID) throw new Error("No test client found");

  const matters = crm_findMattersByClientId(found.CLIENT_ID) || [];
  if (!matters.length) throw new Error("No matters found");

  const docs = crm_listDocumentsByMatterId(matters[0].MATTER_ID, { limit: 20 });
  Logger.log(JSON.stringify(docs, null, 2));
  return docs;
}