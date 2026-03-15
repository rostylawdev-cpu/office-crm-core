/** crmUi.gs */

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu("CRM")
    .addItem("Open Matter Card (active row)", "crm_openMatterCardFromActiveRow")
    .addToUi();
}

function crm_renderMatterCardHtml(matterId) {
  const card = crm_getMatterCard(matterId, {
    tasksStatus: "OPEN",
    limitTasks: 50,
    limitActivities: 20,
  });

  if (!card || !card.ok) {
    return HtmlService.createHtmlOutput(
      `<div style="font-family:Arial;padding:12px">
         <b>Ошибка:</b> ${escapeHtml_(card?.error || "UNKNOWN")}<br/>
         matterId: ${escapeHtml_(matterId || "")}
       </div>`
    ).setTitle("CRM Matter Card");
  }

  const tpl = HtmlService.createTemplateFromFile("crm_matter_card");
  tpl.card = card;
  tpl.matterId = matterId;

  return tpl
    .evaluate()
    .setTitle("CRM Matter Card")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}