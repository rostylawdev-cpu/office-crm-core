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