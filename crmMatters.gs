/** crmMatters.gs */

function crm_createMatter(input) {
  // input: { clientId, category, title, authority, owner? }
  const ss = crm_getSpreadsheet_();
  const c = cfg_();

  if (!input || !input.clientId) throw new Error("crm_createMatter: missing input.clientId");
  if (!input.title) throw new Error("crm_createMatter: missing input.title");

  const sh = ss.getSheetByName(c.SHEETS.MATTERS);
  if (!sh) throw new Error("crm_createMatter: MATTERS sheet not found");

  const now = nowIso_();
  const actor = Session.getActiveUser().getEmail() || "unknown";
  const matterId = makeId_("MAT");

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

  // Activity log
  logActivity_({
    action: "MATTER_CREATED",
    message: `Matter created: ${matterId}`,
    clientId: input.clientId,
    matterId,
    meta: { row: rowIndex, title: rowObj.TITLE, category: rowObj.CATEGORY, authority: rowObj.AUTHORITY }
  });

  // Touch client last activity (если есть такой хедер)
  tryTouchClientLastActivity_(input.clientId, now);

  return { ok: true, matterId, row: rowIndex };
}

function crm_findMattersByClientId(clientId) {
  const ss = crm_getSpreadsheet_();
  const c = cfg_();
  if (!clientId) throw new Error("crm_findMattersByClientId: missing clientId");

  const sh = ss.getSheetByName(c.SHEETS.MATTERS);
  if (!sh) throw new Error("crm_findMattersByClientId: MATTERS sheet not found");

  return findRowsByColumnValue_(sh, c.HEADERS.MATTERS, "CLIENT_ID", clientId);
}

/** ===== helpers (local) ===== */

function makeId_(prefix) {
  // MAT_YYYYMMDD_xxxxxx
  const d = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
  const rnd = Utilities.getUuid().slice(0, 6).toUpperCase();
  return `${prefix}_${d}_${rnd}`;
}

function appendRowByHeaders_(sheet, headers, obj) {
  // headers must be on row 1 already
  const values = headers.map(h => (obj[h] !== undefined ? obj[h] : ""));
  const row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, values.length).setValues([values]);
  return row;
}

function findRowsByColumnValue_(sheet, headers, colName, value) {
  const colIdx = headers.indexOf(colName);
  if (colIdx === -1) throw new Error(`findRowsByColumnValue_: header not found: ${colName}`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const res = [];

  data.forEach((rowVals, i) => {
    if (String(rowVals[colIdx]) === String(value)) {
      res.push(rowToObj_(headers, rowVals, i + 2)); // +2 = реальный номер строки
    }
  });

  return res;
}

function rowToObj_(headers, rowVals, rowNumber) {
  const o = { __row: rowNumber };
  headers.forEach((h, i) => (o[h] = rowVals[i]));
  return o;
}

function logActivity_(p) {
  // p: {action, message, clientId?, matterId?, meta?}
  const ss = crm_getSpreadsheet_();
  const c = cfg_();
  const sh = ss.getSheetByName(c.SHEETS.ACTIVITIES);
  if (!sh) throw new Error("logActivity_: ACTIVITIES sheet not found");

  const actor = Session.getActiveUser().getEmail() || "unknown";
  const now = nowIso_();

  const rowObj = {
    ACTIVITY_ID: makeId_("ACT"),
    TS: now,
    ACTOR: actor,
    CLIENT_ID: p.clientId || "",
    MATTER_ID: p.matterId || "",
    ACTION: p.action || "",
    MESSAGE: p.message || "",
    META_JSON: json_(p.meta || {}),
  };

  appendRowByHeaders_(sh, c.HEADERS.ACTIVITIES, rowObj);
}

function tryTouchClientLastActivity_(clientId, nowIso) {
  const ss = crm_getSpreadsheet_();
  const c = cfg_();
  const sh = ss.getSheetByName(c.SHEETS.CLIENTS);
  if (!sh) return;

  const headers = c.HEADERS.CLIENTS;
  const idxClientId = headers.indexOf("CLIENT_ID");
  const idxLast = headers.indexOf("LAST_ACTIVITY_AT");
  const idxUpdated = headers.indexOf("UPDATED_AT");
  if (idxClientId === -1 || idxLast === -1) return;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const data = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][idxClientId]) === String(clientId)) {
      const rowNumber = i + 2;
      sh.getRange(rowNumber, idxLast + 1).setValue(nowIso);
      if (idxUpdated !== -1) sh.getRange(rowNumber, idxUpdated + 1).setValue(nowIso);
      return;
    }
  }
}

function test_createMatter() {
  // возьми существующий clientId из Clients
  const clientId = "CL_20260304_204602_6MLHYU"; // <-- подставь свой реальный

  const res = crm_createMatter({
    clientId,
    category: "WORK_ACCIDENT",
    title: "Teunat Avoda – test matter",
    authority: "Bituach Leumi",
  });

  Logger.log(res);

  const list = crm_findMattersByClientId(clientId);
  Logger.log(list);
}