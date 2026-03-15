/** crmMatters.gs */

function crm_createMatter(input) {
  // input: { clientId, category, title, authority, owner?, stage?, summaryShort? }
  const ss = crm_getSpreadsheet_();
  const c = cfg_();

  if (!input || !input.clientId) {
    throw new Error("crm_createMatter: missing input.clientId");
  }
  if (!input.title) {
    throw new Error("crm_createMatter: missing input.title");
  }

  const sh = ss.getSheetByName(c.SHEETS.MATTERS);
  if (!sh) throw new Error("crm_createMatter: MATTERS sheet not found");

  const now = nowIso_();
  const actor = getActiveUserEmail_() || "unknown";
  const matterId = generateId_("MAT");

  const rowObj = {
    MATTER_ID: matterId,
    CLIENT_ID: input.clientId,
    CATEGORY: input.category || "",
    TITLE: input.title || "",
    STAGE: input.stage || "NEW",
    OWNER: input.owner || actor,
    AUTHORITY: input.authority || "",
    FOLDER_URL: "",
    OPENED_AT: now,
    CLOSED_AT: "",
    UPDATED_AT: now,
    LAST_ACTIVITY_AT: now,
    SUMMARY_SHORT: input.summaryShort || "",
  };

  const rowIndex = appendRowByHeaders_(sh, c.HEADERS.MATTERS, rowObj);

  crm_logActivity({
    action: "MATTER_CREATED",
    message: `Matter created: ${matterId}`,
    clientId: input.clientId,
    matterId,
    meta: {
      row: rowIndex,
      title: rowObj.TITLE,
      category: rowObj.CATEGORY,
      authority: rowObj.AUTHORITY,
    },
  });

  tryTouchClientLastActivity_(input.clientId, now);

  return { ok: true, matterId, row: rowIndex };
}

function crm_getMatterById(matterId) {
  const ss = crm_getSpreadsheet_();
  const c = cfg_();

  const sh = ss.getSheetByName(c.SHEETS.MATTERS);
  if (!sh) throw new Error("Matters sheet not found");
  if (!matterId) throw new Error("crm_getMatterById: matterId is required");

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return null;

  const header = values[0];
  const rows = values.slice(1);

  const iMatter = header.indexOf("MATTER_ID");
  if (iMatter === -1) throw new Error("MATTER_ID column not found");

  for (let r = 0; r < rows.length; r++) {
    if (String(rows[r][iMatter]) === String(matterId)) {
      return rowToObj_(header, rows[r], r + 2);
    }
  }

  return null;
}

function crm_findMattersByClientId(clientId) {
  const ss = crm_getSpreadsheet_();
  const c = cfg_();

  if (!clientId) throw new Error("crm_findMattersByClientId: missing clientId");

  const sh = ss.getSheetByName(c.SHEETS.MATTERS);
  if (!sh) throw new Error("crm_findMattersByClientId: MATTERS sheet not found");

  return findRowsByColumnValue_(sh, c.HEADERS.MATTERS, "CLIENT_ID", clientId);
}

function crm_createMatterWithTasks(input) {
  // input: { clientId, category, title, authority, stage?, taskTemplateKey?, owner? }

  if (!input || !input.clientId) {
    throw new Error("crm_createMatterWithTasks: missing clientId");
  }
  if (!input.title) {
    throw new Error("crm_createMatterWithTasks: missing title");
  }
  if (!input.category) {
    throw new Error("crm_createMatterWithTasks: missing category");
  }

  // 1) create matter
  const matterRes = crm_createMatter({
    clientId: input.clientId,
    category: input.category,
    title: input.title,
    authority: input.authority || "",
    stage: input.stage || "NEW",
    owner: input.owner || "",
  });

  const matterId = matterRes.matterId;

  // 2) generate tasks from template
  const c = cfg_();
  const key =
    input.taskTemplateKey ||
    (input.category === "WORK_ACCIDENT" ? "WORK_ACCIDENT" : "LABOR_DISPUTE");

  const tpl =
    c.TASK_TEMPLATES && c.TASK_TEMPLATES[key]
      ? c.TASK_TEMPLATES[key]
      : [];

  const createdTasks = [];

  tpl.forEach((t) => {
    const dueDateIso = addDaysIso_(t.days || 0);

    const res = crm_createTask({
      clientId: input.clientId,
      matterId: matterId,
      type: t.type || "GENERAL",
      title: t.title,
      dueDate: dueDateIso,
      status: "OPEN",
      priority: t.priority || "MEDIUM",
      assignee: getActiveUserEmail_() || "unknown",
      generatedBy: "TEMPLATE",
      notes: "",
    });

    createdTasks.push(res);
  });

  // 3) activity log
  crm_logActivity({
    clientId: input.clientId,
    matterId: matterId,
    action: "MATTER_WORKFLOW_CREATED",
    message: `Matter workflow created: ${key}`,
    meta: { templateKey: key, tasks: tpl.length },
  });

  return {
    ok: true,
    matterId,
    tasksCreated: tpl.length,
    taskResults: createdTasks,
  };
}

/** ===== local helper ===== */

function findRowsByColumnValue_(sheet, headers, colName, value) {
  const colIdx = headers.indexOf(colName);
  if (colIdx === -1) {
    throw new Error(`findRowsByColumnValue_: header not found: ${colName}`);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const res = [];

  data.forEach((rowVals, i) => {
    if (String(rowVals[colIdx]) === String(value)) {
      res.push(rowToObj_(headers, rowVals, i + 2));
    }
  });

  return res;
}