/** crmActivities.js */

function crm_logActivity_(ss, payload) {
  const c = cfg_();
  const sh = ss.getSheetByName(c.SHEETS.ACTIVITIES);
  if (!sh) throw new Error("Activities sheet missing");

  const headers = c.HEADERS.ACTIVITIES;

  const rowObj = {
    ACTIVITY_ID: generateId_("ACT"),
    TS: nowIso_(),
    ACTOR: payload.actor ?? getActiveUserEmail_(),
    CLIENT_ID: payload.clientId ?? "",
    MATTER_ID: payload.matterId ?? "",
    ACTION: payload.action ?? "",
    MESSAGE: payload.message ?? "",
    META_JSON: json_(payload.meta ?? {}),
  };

  appendRowByHeaders_(sh, headers, rowObj);
}

function crm_listActivities(filter) {

  filter = filter || {};

  const ss = crm_getSpreadsheet_();
  const c = cfg_();

  const sh = ss.getSheetByName(c.SHEETS.ACTIVITIES);
  if (!sh) throw new Error("Activities sheet not found");

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const header = values[0];
  const rows = values.slice(1);

  const idx = (name) => header.indexOf(name);

  const iClient = idx("CLIENT_ID");

  const clientNeed = filter.clientId ? String(filter.clientId) : "";
  const limit = Number(filter.limit || 20);

  const out = [];

  for (let r = rows.length - 1; r >= 0; r--) {

    const row = rows[r];

    const clientId = String(row[iClient] || "");

    if (clientNeed && clientNeed !== clientId) continue;

    out.push(row);

    if (out.length >= limit) break;
  }

  return out;
}

function crm_findClientById(clientId) {

  const ss = crm_getSpreadsheet_();
  const c = cfg_();

  const sh = ss.getSheetByName(c.SHEETS.CLIENTS);

  const values = sh.getDataRange().getValues();

  const header = values[0];
  const rows = values.slice(1);

  const iClient = header.indexOf("CLIENT_ID");

  for (let r = 0; r < rows.length; r++) {

    if (String(rows[r][iClient]) === String(clientId)) {

      return rows[r];
    }
  }

  return null;
}

function test_dashboard() {

  const clientId = "CL_20260304_204602_6MLHYU";

  const dash = crm_getClientDashboard(clientId);

  Logger.log(JSON.stringify(dash, null, 2));
}