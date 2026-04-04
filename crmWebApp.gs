/** crmWebApp.gs */

function crm_getWebAppDashboard() {
  const c = cfg_();
  const ss = crm_getSpreadsheet_();
  const ssId = ss.getId();
  const sheetNames = ss.getSheets().map(function(s) { return s.getName(); });
  logInfo_("DASHBOARD", "Dashboard load: spreadsheetId=" + ssId + " sheets=[" + sheetNames.join(", ") + "]", {
    spreadsheetId: ssId,
    sheetNames: sheetNames,
  });

  const recentLeads = crm_listLeads({ limit: 10 }) || [];

  const recentClients = crm_getAllRowsFromSheet_(c.SHEETS.CLIENTS, c.HEADERS.CLIENTS)
    .sort(function (a, b) {
      const ta = a.UPDATED_AT ? new Date(a.UPDATED_AT).getTime() : 0;
      const tb = b.UPDATED_AT ? new Date(b.UPDATED_AT).getTime() : 0;
      return tb - ta;
    })
    .slice(0, 10);

  const recentMatters = crm_getAllRowsFromSheet_(c.SHEETS.MATTERS, c.HEADERS.MATTERS)
    .sort(function (a, b) {
      const ta = a.UPDATED_AT ? new Date(a.UPDATED_AT).getTime() : 0;
      const tb = b.UPDATED_AT ? new Date(b.UPDATED_AT).getTime() : 0;
      return tb - ta;
    })
    .slice(0, 10);

  return {
    ok: true,
    stats: {
      leads: crm_getAllRowsFromSheet_(c.SHEETS.LEADS, c.HEADERS.LEADS).length,
      clients: crm_getAllRowsFromSheet_(c.SHEETS.CLIENTS, c.HEADERS.CLIENTS).length,
      matters: crm_getAllRowsFromSheet_(c.SHEETS.MATTERS, c.HEADERS.MATTERS).length,
      tasks: crm_getAllRowsFromSheet_(c.SHEETS.TASKS, c.HEADERS.TASKS).length,
      documents: crm_getAllRowsFromSheet_(c.SHEETS.DOCUMENTS, c.HEADERS.DOCUMENTS).length
    },
    recentLeads: recentLeads,
    recentClients: recentClients,
    recentMatters: recentMatters
  };
}

