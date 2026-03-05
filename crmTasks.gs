/** crmTasks.gs */

function crm_createTask(input) {
  // input: { clientId, matterId?, type?, title, dueDate?, status?, priority?, assignee?, notes?, generatedBy? }
  const ss = crm_getSpreadsheet_();
  const c = cfg_();

  if (!input || !input.clientId) throw new Error("crm_createTask: missing input.clientId");
  if (!input.title) throw new Error("crm_createTask: missing input.title");

  const sh = ss.getSheetByName(c.SHEETS.TASKS);
  if (!sh) throw new Error("crm_createTask: TASKS sheet not found");

  const now = nowIso_();
  const actor = Session.getActiveUser().getEmail() || "unknown";

  const taskId = makeId_("TSK");

  const rowObj = {
    TASK_ID: taskId,
    CLIENT_ID: input.clientId,
    MATTER_ID: input.matterId || "",
    TYPE: input.type || "GENERAL",
    TITLE: input.title,
    DUE_DATE: input.dueDate || "",           // можно ISO или пусто
    STATUS: input.status || "OPEN",          // OPEN / DONE / CANCELED
    PRIORITY: input.priority || "MEDIUM",    // LOW / MEDIUM / HIGH
    GENERATED_BY: input.generatedBy || "MANUAL",
    ASSIGNEE: input.assignee || actor,
    CREATED_AT: now,
    DONE_AT: "",
    NOTES: input.notes || "",
  };

  const rowIndex = appendRowByHeaders_(sh, c.HEADERS.TASKS, rowObj);

  logActivity_({
    action: "TASK_CREATED",
    message: `Task created: ${taskId}`,
    clientId: input.clientId,
    matterId: input.matterId || "",
    meta: { row: rowIndex, title: rowObj.TITLE, due: rowObj.DUE_DATE, priority: rowObj.PRIORITY }
  });

  tryTouchClientLastActivity_(input.clientId, now);

  return { ok: true, taskId, row: rowIndex };
}

function crm_markTaskDone(taskId) {
  const ss = crm_getSpreadsheet_();
  const c = cfg_();

  if (!taskId) throw new Error("crm_markTaskDone: missing taskId");

  const sh = ss.getSheetByName(c.SHEETS.TASKS);
  if (!sh) throw new Error("crm_markTaskDone: TASKS sheet not found");

  const headers = c.HEADERS.TASKS;
  const idxTaskId = headers.indexOf("TASK_ID");
  const idxStatus = headers.indexOf("STATUS");
  const idxDoneAt = headers.indexOf("DONE_AT");
  const idxUpdatedAt = headers.indexOf("UPDATED_AT"); // может не быть, ок
  const idxClientId = headers.indexOf("CLIENT_ID");
  const idxMatterId = headers.indexOf("MATTER_ID");

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: false, reason: "no rows" };

  const data = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const now = nowIso_();

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][idxTaskId]) === String(taskId)) {
      const rowNumber = i + 2;

      if (idxStatus !== -1) sh.getRange(rowNumber, idxStatus + 1).setValue("DONE");
      if (idxDoneAt !== -1) sh.getRange(rowNumber, idxDoneAt + 1).setValue(now);
      if (idxUpdatedAt !== -1) sh.getRange(rowNumber, idxUpdatedAt + 1).setValue(now);

      const clientId = idxClientId !== -1 ? data[i][idxClientId] : "";
      const matterId = idxMatterId !== -1 ? data[i][idxMatterId] : "";

      logActivity_({
        action: "TASK_DONE",
        message: `Task done: ${taskId}`,
        clientId: clientId || "",
        matterId: matterId || "",
        meta: { row: rowNumber }
      });

      if (clientId) tryTouchClientLastActivity_(clientId, now);

      return { ok: true, taskId, row: rowNumber };
    }
  }

  return { ok: false, reason: "task not found" };
}

function crm_findOpenTasksByClientId(clientId) {
  const ss = crm_getSpreadsheet_();
  const c = cfg_();

  if (!clientId) throw new Error("crm_findOpenTasksByClientId: missing clientId");

  const sh = ss.getSheetByName(c.SHEETS.TASKS);
  if (!sh) throw new Error("crm_findOpenTasksByClientId: TASKS sheet not found");

  const headers = c.HEADERS.TASKS;
  const idxClientId = headers.indexOf("CLIENT_ID");
  const idxStatus = headers.indexOf("STATUS");
  if (idxClientId === -1) throw new Error("TASKS missing CLIENT_ID header");
  if (idxStatus === -1) throw new Error("TASKS missing STATUS header");

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const data = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const res = [];

  data.forEach((rowVals, i) => {
    const okClient = String(rowVals[idxClientId]) === String(clientId);
    const okStatus = String(rowVals[idxStatus]).toUpperCase() === "OPEN";
    if (okClient && okStatus) res.push(rowToObj_(headers, rowVals, i + 2));
  });

  return res;
}

function test_createTask() {
  const clientId = "CL_20260304_204602_6MLHYU"; // <-- свой
  const matterId = "MAT_20260305_......";       // <-- из Matters

  const res = crm_createTask({
    clientId,
    matterId,
    type: "FOLLOW_UP",
    title: "Позвонить клиенту и запросить טופס 250",
    dueDate: nowIso_(),
    priority: "HIGH",
    notes: "Тестовая задача",
  });

  Logger.log(res);

  const open = crm_findOpenTasksByClientId(clientId);
  Logger.log(open);
}

/**
 * List tasks by filters
 * filter: { clientId?, matterId?, assignee?, status?="OPEN", limit?=50 }
 */
function crm_listTasks(filter) {
  filter = filter || {};
  const ss = crm_getSpreadsheet_();
  const c = cfg_();
  const sh = ss.getSheetByName(c.SHEETS.TASKS);
  if (!sh) throw new Error("crm_listTasks: TASKS sheet not found");

  const statusNeed = (filter.status || "OPEN").toString();
  const clientNeed = filter.clientId ? String(filter.clientId) : "";
  const matterNeed = filter.matterId ? String(filter.matterId) : "";
  const assigneeNeed = filter.assignee ? String(filter.assignee) : "";
  const limit = Number(filter.limit || 50);

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const header = values[0];
  const rows = values.slice(1);

  const idx = (name) => header.indexOf(name);

  const iTaskId = idx("TASK_ID");
  const iClientId = idx("CLIENT_ID");
  const iMatterId = idx("MATTER_ID");
  const iTitle = idx("TITLE");
  const iDue = idx("DUE_DATE");
  const iStatus = idx("STATUS");
  const iPriority = idx("PRIORITY");
  const iAssignee = idx("ASSIGNEE");
  const iCreated = idx("CREATED_AT");
  const iDone = idx("DONE_AT");
  const iNotes = idx("NOTES");
  const iType = idx("TYPE");
  const iGeneratedBy = idx("GENERATED_BY");

  const out = [];
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.join("") === "") continue;

    const status = String(row[iStatus] || "");
    if (statusNeed && status !== statusNeed) continue;

    const clientId = String(row[iClientId] || "");
    if (clientNeed && clientId !== clientNeed) continue;

    const matterId = String(row[iMatterId] || "");
    if (matterNeed && matterId !== matterNeed) continue;

    const assignee = String(row[iAssignee] || "");
    if (assigneeNeed && assignee !== assigneeNeed) continue;

    out.push({
      row: r + 2, // sheet row number (1-based + header)
      taskId: row[iTaskId],
      clientId,
      matterId,
      type: row[iType],
      title: row[iTitle],
      dueDate: row[iDue],
      status,
      priority: row[iPriority],
      assignee,
      createdAt: row[iCreated],
      doneAt: row[iDone],
      generatedBy: row[iGeneratedBy],
      notes: row[iNotes],
    });

    if (out.length >= limit) break;
  }

  // Сортировка: сначала DUE_DATE (пустые вниз), потом PRIORITY
  const prRank = { HIGH: 1, MEDIUM: 2, LOW: 3 };
  out.sort((a, b) => {
    const da = a.dueDate ? new Date(a.dueDate).getTime() : Number.POSITIVE_INFINITY;
    const db = b.dueDate ? new Date(b.dueDate).getTime() : Number.POSITIVE_INFINITY;
    if (da !== db) return da - db;
    const pa = prRank[String(a.priority || "").toUpperCase()] || 9;
    const pb = prRank[String(b.priority || "").toUpperCase()] || 9;
    return pa - pb;
  });

  return out;
}

/**
 * Mark task as DONE by taskId
 * @param {string} taskId
 * @param {string=} note optional note appended to NOTES
 * @returns {{ok:boolean, taskId:string, row:number}}
 */
function crm_markTaskDone(taskId, note) {
  if (!taskId) throw new Error("crm_markTaskDone: missing taskId");

  const ss = crm_getSpreadsheet_();
  const c = cfg_();
  const sh = ss.getSheetByName(c.SHEETS.TASKS);
  if (!sh) throw new Error("crm_markTaskDone: TASKS sheet not found");

  const values = sh.getDataRange().getValues();
  if (values.length < 2) throw new Error("crm_markTaskDone: no data in TASKS");

  const header = values[0];
  const idx = (name) => header.indexOf(name);

  const iTaskId = idx("TASK_ID");
  const iStatus = idx("STATUS");
  const iDoneAt = idx("DONE_AT");
  const iUpdatedAt = idx("UPDATED_AT"); // может отсутствовать — ок
  const iNotes = idx("NOTES");
  const iClientId = idx("CLIENT_ID");
  const iMatterId = idx("MATTER_ID");
  const iTitle = idx("TITLE");

  if (iTaskId < 0 || iStatus < 0 || iDoneAt < 0) {
    throw new Error("crm_markTaskDone: required columns missing (TASK_ID/STATUS/DONE_AT)");
  }

  const now = nowIso_();

  // find row
  let rowNum = -1;
  let row = null;
  for (let r = 1; r < values.length; r++) {
    const curId = String(values[r][iTaskId] || "");
    if (curId === String(taskId)) {
      rowNum = r + 1; // sheet row number (1-based)
      row = values[r];
      break;
    }
  }
  if (rowNum < 0) throw new Error("crm_markTaskDone: task not found: " + taskId);

  // update cells
  sh.getRange(rowNum, iStatus + 1).setValue("DONE");
  sh.getRange(rowNum, iDoneAt + 1).setValue(now);
  if (iUpdatedAt >= 0) sh.getRange(rowNum, iUpdatedAt + 1).setValue(now);

  if (iNotes >= 0 && note) {
    const prev = String(row[iNotes] || "");
    const next = prev ? (prev + "\n" + note) : String(note);
    sh.getRange(rowNum, iNotes + 1).setValue(next);
  }

  // activity log
  const actor = Session.getActiveUser().getEmail() || "unknown";
  const clientId = iClientId >= 0 ? String(row[iClientId] || "") : "";
  const matterId = iMatterId >= 0 ? String(row[iMatterId] || "") : "";
  const title = iTitle >= 0 ? String(row[iTitle] || "") : "";

  logInfo_("TASK_DONE", "Task marked DONE: " + taskId, {
    taskId,
    row: rowNum,
    title,
    clientId,
    matterId,
    actor,
  });

  return { ok: true, taskId: String(taskId), row: rowNum };
}

function crm_reopenTask(taskId) {
  if (!taskId) throw new Error("crm_reopenTask: missing taskId");

  const ss = crm_getSpreadsheet_();
  const c = cfg_();
  const sh = ss.getSheetByName(c.SHEETS.TASKS);
  if (!sh) throw new Error("crm_reopenTask: TASKS sheet not found");

  const values = sh.getDataRange().getValues();
  const header = values[0];
  const idx = (name) => header.indexOf(name);

  const iTaskId = idx("TASK_ID");
  const iStatus = idx("STATUS");
  const iDoneAt = idx("DONE_AT");
  const iUpdatedAt = idx("UPDATED_AT");

  if (iTaskId < 0 || iStatus < 0 || iDoneAt < 0) {
    throw new Error("crm_reopenTask: required columns missing (TASK_ID/STATUS/DONE_AT)");
  }

  let rowNum = -1;
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][iTaskId] || "") === String(taskId)) {
      rowNum = r + 1;
      break;
    }
  }
  if (rowNum < 0) throw new Error("crm_reopenTask: task not found: " + taskId);

  const now = nowIso_();
  sh.getRange(rowNum, iStatus + 1).setValue("OPEN");
  sh.getRange(rowNum, iDoneAt + 1).setValue("");
  if (iUpdatedAt >= 0) sh.getRange(rowNum, iUpdatedAt + 1).setValue(now);

  logInfo_("TASK_REOPEN", "Task reopened: " + taskId, { taskId, row: rowNum });
  return { ok: true, taskId: String(taskId), row: rowNum };
}

//test function test_markDone
function test_markDone() {
  const taskId = "TSK_20260305_2A70B1"; // подставь свой
  const res = crm_markTaskDone(taskId, "Закрыто автотестом");
  Logger.log(JSON.stringify(res, null, 2));
  return res;
}