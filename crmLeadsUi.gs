/** crmLeadsUi.gs */

function crm_openLeadCardById(leadId) {
  if (!leadId) {
    SpreadsheetApp.getUi().alert("crm_openLeadCardById: missing leadId");
    return;
  }
  crm_openLeadCard_(leadId);
}

function crm_openLeadCardFromActiveRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  const row = sh.getActiveRange().getRow();

  const leadId = getRowValueByHeader_(sh, row, "LEAD_ID");

  if (!leadId) {
    SpreadsheetApp.getUi().alert("Не удалось найти LEAD_ID в активной строке.");
    return;
  }

  crm_openLeadCard_(String(leadId).trim());
}

function crm_openLeadCard_ (leadId) {
  const card = crm_getLeadCard(leadId, {
    limitActivities: 50,
  });

  if (!card || !card.ok) {
    SpreadsheetApp.getUi().alert(`Лид не найден: ${leadId}`);
    return;
  }

  const t = HtmlService.createTemplateFromFile("crm_lead_card");
  t.card = card;
  t.leadId = leadId;

  const html = t
    .evaluate()
    .setTitle("CRM Lead Card")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  SpreadsheetApp.getUi().showSidebar(html);
}

function crm_renderLeadCardHtml(leadId) {
  const card = crm_getLeadCard(leadId, {
    limitActivities: 20,
  });

  if (!card || !card.ok) {
    return HtmlService.createHtmlOutput(
      `<div style="font-family:Arial;padding:12px">
         <b>Ошибка:</b> ${escapeHtml_(card?.error || "UNKNOWN")}<br/>
         leadId: ${escapeHtml_(leadId || "")}
       </div>`
    ).setTitle("CRM Lead Card");
  }

  const tpl = HtmlService.createTemplateFromFile("crm_lead_card");
  tpl.card = card;
  tpl.leadId = leadId;

  return tpl
    .evaluate()
    .setTitle("CRM Lead Card")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Action из HTML: convert lead -> client
 */
function crm_uiConvertLeadToClient(leadId) {
  if (!leadId) throw new Error("crm_uiConvertLeadToClient: missing leadId");
  return crm_convertLeadToClient(leadId);
}

/**
 * Action из HTML: convert lead -> client + matter
 * Сейчас делаем дефолтный matter из лида.
 */
function crm_uiConvertLeadToClientAndMatter(leadId) {
  if (!leadId) throw new Error("crm_uiConvertLeadToClientAndMatter: missing leadId");

  const lead = crm_getLeadById(leadId);
  if (!lead) throw new Error("Lead not found: " + leadId);

  const caseType = String(lead.CASE_TYPE || "").trim().toUpperCase();
  const isWorkAccident = caseType === "WORK_ACCIDENT";

  return crm_convertLeadToClientAndMatter(leadId, {
    category: caseType || "GENERAL",
    title: `Matter from lead: ${lead.FULL_NAME || "Untitled"}`,
    authority: isWorkAccident ? "Bituach Leumi" : "",
    stage: "NEW",
    taskTemplateKey: isWorkAccident ? "WORK_ACCIDENT" : "",
    owner: lead.ASSIGNED_TO || getActiveUserEmail_() || "",
  });
}

/**
 * Signing-first onboarding context setup.
 * Provisions a client record and onboarding matter for the lead WITHOUT
 * marking the lead as CONVERTED. Conversion is deferred until signing success.
 * Called by crm_webStartLeadOnboarding.
 */
function crm_initLeadOnboardingContext_(leadId) {
  const lead = crm_getLeadById(leadId);
  if (!lead) throw new Error("crm_initLeadOnboardingContext_: lead not found: " + leadId);

  const currentStatus = String(lead.STATUS || "").toUpperCase();

  // If already fully converted, try to find the ONBOARDING matter specifically.
  // Never fall back to an arbitrary "latest" matter — that could be any unrelated case.
  // If no onboarding matter exists, return matterId:null so the UI can land on the client.
  if (currentStatus === "CONVERTED" && lead.CLIENT_ID) {
    const convertedClientId = String(lead.CLIENT_ID).trim();
    var mattersForConverted = crm_findMattersByClientId(convertedClientId) || [];
    // Fallback: if matter.CLIENT_ID is still provisional (reattachment was interrupted),
    // search by leadId so repeated Start Onboarding still finds the right matter.
    if (mattersForConverted.length === 0) {
      mattersForConverted = crm_findMattersByClientId(leadId) || [];
    }
    const onboardingMatter = mattersForConverted.find(function(m) {
      return (m.STAGE || "").toUpperCase() === "ONBOARDING";
    });
    return {
      leadId,
      clientId: convertedClientId,
      matterId: onboardingMatter ? onboardingMatter.MATTER_ID : null,
      alreadyConverted: true,
    };
  }

  // No client pre-creation: real client is created ONLY after successful onboarding signing.
  // leadId is used as provisional CLIENT_ID in matter + document records until crm_finalizeOnboardingConversion_ runs.

  // Set lead status to ONBOARDING (defers CONVERTED until post-sign)
  if (currentStatus !== "ONBOARDING" && currentStatus !== "PENDING_SIGNATURE") {
    crm_updateLeadStatus(leadId, "ONBOARDING");
  }

  // Idempotency: reuse existing ONBOARDING matter anchored to this leadId.
  // If somehow multiple were created (data drift/past bugs), pick the newest and warn.
  const existingMatters = crm_findMattersByClientId(leadId) || [];
  const allOnboardingMatters = existingMatters.filter(function(m) {
    return (m.STAGE || "").toUpperCase() === "ONBOARDING";
  });
  if (allOnboardingMatters.length > 0) {
    // Sort by OPENED_AT descending — newest first
    allOnboardingMatters.sort(function(a, b) {
      return (b.OPENED_AT || "") > (a.OPENED_AT || "") ? 1 : -1;
    });
    const chosenMatter = allOnboardingMatters[0];
    if (allOnboardingMatters.length > 1) {
      crm_logActivity({
        action: "ONBOARDING_INVARIANT_WARNING",
        message: "Multiple ONBOARDING matters found for lead " + leadId + " — reusing newest: " + chosenMatter.MATTER_ID,
        matterId: chosenMatter.MATTER_ID,
        meta: { leadId: leadId, count: allOnboardingMatters.length, chosenMatterId: chosenMatter.MATTER_ID },
      });
    }
    return { leadId, clientId: null, matterId: chosenMatter.MATTER_ID, reused: true };
  }

  // No onboarding matter yet — create one with leadId as provisional CLIENT_ID
  const caseType = String(lead.CASE_TYPE || "").trim().toUpperCase();
  const categoryLabels_ = {
    WORK_ACCIDENT: "Work Accident",
    LABOR: "Labor",
    LABOR_DISPUTE: "Labor Dispute",
    EMPLOYMENT: "Employment",
    GENERAL: "General",
  };
  const categoryLabel_ = categoryLabels_[caseType] || (caseType || "General");
  const dateStr_ = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy");
  // Prefer lead's bilingual subject as the matter title; fall back to category + date
  const onboardingTitle_ = String(lead.SUBJECT_RU || lead.SUBJECT_HE || "").trim()
    || (categoryLabel_ + " – " + dateStr_);
  const matterRes = crm_createMatter({
    clientId: leadId,  // provisional — real client created by crm_finalizeOnboardingConversion_ after signing
    category: caseType || "GENERAL",
    title: onboardingTitle_,
    subjectRu: String(lead.SUBJECT_RU || "").trim(),
    subjectHe: String(lead.SUBJECT_HE || "").trim(),
    authority: caseType === "WORK_ACCIDENT" ? "Bituach Leumi" : "",
    stage: "ONBOARDING",
    owner: String(lead.ASSIGNED_TO || getActiveUserEmail_() || "").trim(),
  });

  crm_logActivity({
    action: "LEAD_ONBOARDING_STARTED",
    message: "Signing-first onboarding initiated for lead " + leadId,
    matterId: matterRes.matterId,
    meta: { leadId, matterId: matterRes.matterId },
  });

  return { leadId, clientId: null, matterId: matterRes.matterId };
}