/** crmSearchUi.gs */

function crm_openSearchDialog() {
  const tpl = HtmlService.createTemplateFromFile("crm_search");
  const html = tpl
    .evaluate()
    .setTitle("CRM Search")
    .setWidth(900)
    .setHeight(700);

  SpreadsheetApp.getUi().showModelessDialog(html, "CRM Search");
}

function crm_runSearch(query) {
  const res = crm_searchAll(query, { limit: 20 });
  // Sanitize all row arrays — Date objects in sheet cells cause google.script.run to return null silently.
  return {
    ok: res.ok,
    query: res.query,
    leads:   crm_sanitizeRowsForClient_(res.leads   || []),
    clients: crm_sanitizeRowsForClient_(res.clients || []),
    matters: crm_sanitizeRowsForClient_(res.matters || []),
    meta: res.meta,
  };
}