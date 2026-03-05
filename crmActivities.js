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