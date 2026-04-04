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
  return crm_searchAll(query, { limit: 20 });
}