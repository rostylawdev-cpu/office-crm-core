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

  // Load each sheet once; reuse for both counts and recent-item lists
  const allLeads     = crm_getAllRowsFromSheet_(c.SHEETS.LEADS,     c.HEADERS.LEADS);
  const allClients   = crm_getAllRowsFromSheet_(c.SHEETS.CLIENTS,   c.HEADERS.CLIENTS);
  const allMatters   = crm_getAllRowsFromSheet_(c.SHEETS.MATTERS,   c.HEADERS.MATTERS);
  const allTasks     = crm_getAllRowsFromSheet_(c.SHEETS.TASKS,     c.HEADERS.TASKS);
  const allDocuments = crm_getAllRowsFromSheet_(c.SHEETS.DOCUMENTS, c.HEADERS.DOCUMENTS);

  logInfo_("DASHBOARD", "Row counts: leads=" + allLeads.length + " clients=" + allClients.length +
    " matters=" + allMatters.length + " tasks=" + allTasks.length + " docs=" + allDocuments.length, {});

  function byUpdatedDesc(a, b) {
    const ta = a.UPDATED_AT ? new Date(a.UPDATED_AT).getTime() : 0;
    const tb = b.UPDATED_AT ? new Date(b.UPDATED_AT).getTime() : 0;
    return tb - ta;
  }

  const recentLeads   = allLeads.slice().sort(byUpdatedDesc).slice(0, 10);
  const recentClients = allClients.slice().sort(byUpdatedDesc).slice(0, 10);
  const recentMatters = allMatters.slice().sort(byUpdatedDesc).slice(0, 10);

  return {
    ok: true,
    stats: {
      leads:     Number(allLeads.length),
      clients:   Number(allClients.length),
      matters:   Number(allMatters.length),
      tasks:     Number(allTasks.length),
      documents: Number(allDocuments.length),
    },
    // Sanitize row objects so google.script.run can serialize them:
    // Date cell values become ISO strings; Error/unknown types become "".
    recentLeads:   crm_sanitizeRowsForClient_(recentLeads),
    recentClients: crm_sanitizeRowsForClient_(recentClients),
    recentMatters: crm_sanitizeRowsForClient_(recentMatters),
  };
}

// Convert row objects from sheet data into plain JSON-safe objects.
// google.script.run can fail silently (returning null) when the return value
// contains non-serializable types such as raw Date objects or formula error values.
function crm_sanitizeRowsForClient_(rows) {
  return (rows || []).map(function(row) {
    var out = {};
    for (var k in row) {
      if (k === '__row') continue; // internal metadata, not needed by client
      var v = row[k];
      if (v === null || v === undefined || v === '') {
        out[k] = '';
      } else if (v instanceof Date) {
        out[k] = isNaN(v.getTime()) ? '' : v.toISOString();
      } else if (typeof v === 'number' || typeof v === 'boolean') {
        out[k] = v;
      } else if (typeof v === 'string') {
        out[k] = v;
      } else {
        // Formula error values (#N/A, #VALUE!, etc.) and any other non-primitive
        try { out[k] = String(v); } catch (e) { out[k] = ''; }
      }
    }
    return out;
  });
}

