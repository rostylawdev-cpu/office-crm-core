function crm_getMatterCard(matterId, opt) {
  opt = opt || {};
  const limitTasks = opt.limitTasks ?? 20;
  const limitActivities = opt.limitActivities ?? 20;
  const tasksStatus = opt.tasksStatus ?? "OPEN"; // или null = все

  const ss = crm_getSpreadsheet_();
  const c = cfg_();

  if (!matterId) throw new Error("crm_getMatterCard: missing matterId");

  // Matter
  const matter = crm_getMatterById_(ss, matterId);
  if (!matter) return { ok: false, error: "MATTER_NOT_FOUND", matterId };

  // Client
  const client = crm_getClientById_(ss, matter.CLIENT_ID);

  // Tasks
  const tasks = crm_listTasksByMatter_(ss, matterId, {
    status: tasksStatus,
    limit: limitTasks
  });

  // Activities
  const activities = crm_listActivities_(ss, {
    matterId,
    limit: limitActivities
  });

  return { ok: true, client, matter, tasks, activities };
}

function crm_getMatterById_(ss, matterId) {
  const c = cfg_();
  const sh = ss.getSheetByName(c.SHEETS.MATTERS);
  const rows = sh.getDataRange().getValues();
  const headers = rows[0];
  const idx = headers.indexOf("MATTER_ID");
  if (idx < 0) throw new Error("MATTERS missing MATTER_ID header");

  for (let r = 1; r < rows.length; r++) {
    if (rows[r][idx] === matterId) return rowToObj_(headers, rows[r], r + 1);
  }
  return null;
}

function crm_getClientById_(ss, clientId) {
  const c = cfg_();
  const sh = ss.getSheetByName(c.SHEETS.CLIENTS);
  const rows = sh.getDataRange().getValues();
  const headers = rows[0];
  const idx = headers.indexOf("CLIENT_ID");
  if (idx < 0) throw new Error("CLIENTS missing CLIENT_ID header");

  for (let r = 1; r < rows.length; r++) {
    if (rows[r][idx] === clientId) return rowToObj_(headers, rows[r], r + 1);
  }
  return null;
}

function crm_listTasksByMatter_(ss, matterId, opt) {
  opt = opt || {};
  const c = cfg_();
  const sh = ss.getSheetByName(c.SHEETS.TASKS);

  const rows = sh.getDataRange().getValues();
  const headers = rows[0];
  const iMatter = headers.indexOf("MATTER_ID");
  const iStatus = headers.indexOf("STATUS");
  if (iMatter < 0) throw new Error("TASKS missing MATTER_ID");
  if (iStatus < 0) throw new Error("TASKS missing STATUS");

  const out = [];
  for (let r = 1; r < rows.length; r++) {
    if (rows[r][iMatter] !== matterId) continue;
    if (opt.status && rows[r][iStatus] !== opt.status) continue;
    out.push(rowToObj_(headers, rows[r], r + 1));
  }

  // можно позже сделать сортировку по DUE_DATE/CREATED_AT
  return out.slice(0, opt.limit ?? 50);
}

function crm_listActivities_(ss, opt) {
  opt = opt || {};
  const c = cfg_();
  const sh = ss.getSheetByName(c.SHEETS.ACTIVITIES);

  const rows = sh.getDataRange().getValues();
  const headers = rows[0];
  const iMatter = headers.indexOf("MATTER_ID");
  const iClient = headers.indexOf("CLIENT_ID");
  if (iMatter < 0) throw new Error("ACTIVITIES missing MATTER_ID");
  if (iClient < 0) throw new Error("ACTIVITIES missing CLIENT_ID");

  const out = [];
  for (let r = rows.length - 1; r >= 1; r--) { // с конца: “последние сверху”
    const mOk = opt.matterId ? rows[r][iMatter] === opt.matterId : true;
    const cOk = opt.clientId ? rows[r][iClient] === opt.clientId : true;
    if (!mOk || !cOk) continue;
    out.push(rowToObj_(headers, rows[r], r + 1));
    if (out.length >= (opt.limit ?? 20)) break;
  }
  return out;
}

function test_getMatterCard() {
  const card = crm_getMatterCard("MAT_20260305_803005", {
    tasksStatus: "OPEN",
    limitTasks: 50,
    limitActivities: 20
  });
  Logger.log(JSON.stringify(card, null, 2));
}