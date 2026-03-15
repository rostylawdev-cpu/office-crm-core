/** crmActivities.gs */

function crm_logActivity(payload) {
  const ss = crm_getSpreadsheet_();
  return crm_logActivity_(ss, payload);
}

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
  const iMatter = idx("MATTER_ID");

  const clientNeed = filter.clientId ? String(filter.clientId) : "";
  const matterNeed = filter.matterId ? String(filter.matterId) : "";
  const limit = Number(filter.limit || 20);

  const out = [];

  for (let r = rows.length - 1; r >= 0; r--) {
    const row = rows[r];

    const clientId = String(row[iClient] || "");
    const matterId = String(row[iMatter] || "");

    if (clientNeed && clientNeed !== clientId) continue;
    if (matterNeed && matterNeed !== matterId) continue;

    out.push(rowToObj_(header, row, r + 2));

    if (out.length >= limit) break;
  }

  return out;
}