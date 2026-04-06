/** crmUi.gs */

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu("CRM")
    .addItem("Open Matter Card (active row)", "crm_openMatterCardFromActiveRow")
    .addItem("Open Lead Card (active row)", "crm_openLeadCardFromActiveRow")
    .addItem("Open Client Card (active row)", "crm_openClientCardFromActiveRow")
    .addSeparator()
    .addItem("Search CRM", "crm_openSearchDialog")
    .addSeparator()
    .addItem("Open CRM Web App", "crm_openWebAppHome")
    .addSeparator()
    .addItem("Run Consistency Audit", "crm_runConsistencyAudit")
    .addItem("Run Consistency Repair", "crm_runConsistencyRepair")
    .addToUi();
}

function crm_openWebAppHome() {
  const url = ScriptApp.getService().getUrl();
  if (!url) {
    SpreadsheetApp.getUi().alert("Deploy the script as Web App first.");
    return;
  }

  const target = url + "?page=home&ts=" + Date.now();

  const html = HtmlService.createHtmlOutput(
    '<script>window.open("' + target + '", "_blank");google.script.host.close();</script>'
  );
  SpreadsheetApp.getUi().showModalDialog(html, "Open CRM Web App");
}

function doGet(e) {
  try {
    const rawId = (e && e.parameter && e.parameter.id)
      ? String(e.parameter.id).trim()
      : "";
    const leadId = (e && e.parameter && e.parameter.leadId)
      ? String(e.parameter.leadId).trim()
      : "";
    const id = rawId || leadId;

    const page = (e && e.parameter && e.parameter.page)
      ? String(e.parameter.page).trim()
      : "home";

    const webAppBase = ScriptApp.getService().getUrl();
    if (!webAppBase) {
      throw new Error("Web App URL unavailable. Deploy as Web App first.");
    }

    if (page === "lead") {
      if (!id) return crm_renderErrorPage_("Missing lead id", { page, id });

      const card = crm_getLeadCard(id, { limitActivities: 20 });
      const tpl = HtmlService.createTemplateFromFile("crm_app_lead");
      tpl.card = card;
      tpl.entityId = id;
      tpl.page = page;
      tpl.webAppBase = webAppBase;

      return tpl.evaluate()
        .setTitle("CRM Lead")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (page === "client") {
      if (!id) return crm_renderErrorPage_("Missing client id", { page, id });

      const card = crm_getClientCard(id, {
        tasksStatus: "OPEN",
        limitTasks: 20,
        limitActivities: 20,
        limitDocuments: 20,
      });

      const tpl = HtmlService.createTemplateFromFile("crm_app_client");
      tpl.card = card;
      tpl.entityId = id;
      tpl.page = page;
      tpl.webAppBase = webAppBase;

      return tpl.evaluate()
        .setTitle("CRM Client")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (page === "matter") {
      if (!id) return crm_renderErrorPage_("Missing matter id", { page, id });

      const card = crm_getMatterCard(id, {
        tasksStatus: "OPEN",
        limitTasks: 20,
        limitActivities: 20,
        limitDocuments: 20,
      });

      const tpl = HtmlService.createTemplateFromFile("crm_app_matter");
      tpl.card = card;
      tpl.entityId = id;
      tpl.page = page;
      tpl.webAppBase = webAppBase;

      return tpl.evaluate()
        .setTitle("CRM Matter")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (page === "search") {
      const tpl = HtmlService.createTemplateFromFile("crm_app_search");
      tpl.page = page;
      tpl.entityId = id;
      tpl.webAppBase = webAppBase;

      return tpl.evaluate()
        .setTitle("CRM Search")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (page === "sign") {
      const token = (e && e.parameter && e.parameter.token) ? String(e.parameter.token).trim() : "";
      if (!token) return crm_renderErrorPage_("Missing token for sign page", { page, token });

      const tpl = HtmlService.createTemplateFromFile("crm_app_sign");
      tpl.page = page;
      tpl.token = token;
      tpl.webAppBase = webAppBase;

      return tpl.evaluate()
        .setTitle("CRM Sign Document")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (page === "new_lead") {
      const tpl = HtmlService.createTemplateFromFile("crm_app_new_lead");
      tpl.page = page;
      tpl.entityId = id;
      tpl.webAppBase = webAppBase;

      return tpl.evaluate()
        .setTitle("CRM New Lead")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (page === "start_onboarding") {
      if (!id) return crm_renderErrorPage_("Missing lead id for onboarding", { page, id });
      const onboardResult = crm_webStartLeadOnboarding(id);
      // No real client exists before signing — fall back to lead page as backup URL
      const clientUrl = (onboardResult.clientId && !String(onboardResult.clientId || "").startsWith("LEAD_"))
        ? webAppBase + "?page=client&id=" + encodeURIComponent(onboardResult.clientId)
        : webAppBase + "?page=lead&id=" + encodeURIComponent(id);
      // If no onboarding matter exists (e.g. already converted without one), land on client
      const matterUrl = onboardResult.matterId
        ? webAppBase + "?page=matter&id=" + encodeURIComponent(onboardResult.matterId)
        : null;
      const redirectUrl = matterUrl || clientUrl;
      const matterBtn = matterUrl
        ? '<a class="btn" href="' + crm_escapeHtml_(matterUrl) + '">Open Onboarding Matter</a>'
        : '';
      return HtmlService.createHtmlOutput(
        '<!doctype html><html><head><meta charset="utf-8">' +
        '<style>' +
        'body{font-family:Arial,sans-serif;background:#f6f8fc;padding:32px;color:#1f2937;}' +
        '.box{background:white;border:1px solid #e5e7eb;border-radius:16px;padding:28px;max-width:480px;margin:0 auto;}' +
        'h2{margin-top:0;font-size:20px;}p{color:#6b7280;}' +
        '.btn{display:inline-block;text-decoration:none;padding:10px 18px;border-radius:10px;' +
        'background:#1a73e8;color:white;font-weight:600;margin:8px 8px 0 0;}' +
        '.btn.sec{background:#6b7280;}' +
        '</style>' +
        '<script>setTimeout(function(){window.location.href=' + JSON.stringify(redirectUrl) + ';},800);</script>' +
        '</head><body>' +
        '<div class="box">' +
        '<h2>Onboarding Ready</h2>' +
        '<p>Onboarding matter created for lead <b>' + crm_escapeHtml_(id) + '</b>. No client record is created until signing is complete.</p>' +
        '<p style="background:#fef3c7;border-left:3px solid #f59e0b;padding:8px 12px;border-radius:0 8px 8px 0;font-size:13px;color:#92400e;margin:12px 0;">' +
        '\u26a0\ufe0f The lead is <b>not yet converted.</b> Conversion happens automatically after both documents are signed.' +
        '</p>' +
        '<ol style="font-size:14px;padding-left:20px;line-height:2;margin:12px 0;">' +
        '<li>Open the onboarding matter (button below)</li>' +
        '<li>Click <b>Generate Agreement + POA</b></li>' +
        '<li>Click <b>Create Sign Links</b> and send to client for signing</li>' +
        '<li>After both documents are signed, the lead converts automatically</li>' +
        '</ol>' +
        '<p style="color:#6b7280;font-size:12px;">Redirecting to matter in 1 second\u2026</p>' +
        matterBtn +
        '<a class="btn sec" href="' + crm_escapeHtml_(clientUrl) + '">' + (onboardResult.clientId && !String(onboardResult.clientId || "").startsWith("LEAD_") ? "Open Client" : "Back to Lead") + '</a>' +
        '</div></body></html>'
      ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (page === "new_client") {
      const tpl = HtmlService.createTemplateFromFile("crm_app_new_client");
      tpl.page = page;
      tpl.entityId = id;
      tpl.webAppBase = webAppBase;

      return tpl.evaluate()
        .setTitle("CRM New Client")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (page === "new_matter") {
      const tpl = HtmlService.createTemplateFromFile("crm_app_new_matter");
      tpl.page = page;
      tpl.entityId = id;
      tpl.webAppBase = webAppBase;

      return tpl.evaluate()
        .setTitle("CRM New Matter")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (page === "edit_client") {
      const clientId = (e && e.parameter && (e.parameter.clientId || e.parameter.id))
        ? String(e.parameter.clientId || e.parameter.id).trim()
        : "";

      if (!clientId) return crm_renderErrorPage_("Missing clientId for edit", { page, clientId });

      const client = crm_findClientById(clientId);
      if (!client) return crm_renderErrorPage_("Client not found: " + clientId, { page, clientId });

      const tpl = HtmlService.createTemplateFromFile("crm_app_edit_client");
      tpl.page = page;
      tpl.clientId = clientId;
      tpl.client = client;
      tpl.webAppBase = webAppBase;

      return tpl.evaluate()
        .setTitle("CRM Edit Client")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (page === "new_task") {
      const rawMatterId = (e && e.parameter && (e.parameter.matterId || e.parameter.id))
        ? String(e.parameter.matterId || e.parameter.id).trim()
        : "";

      const matterId = rawMatterId;

      if (!matterId) return crm_renderErrorPage_("Missing matterId for task creation", { page, matterId });

      if (matterId.startsWith("CL_") || matterId.startsWith("CLIENT_")) {
        return crm_renderErrorPage_("new_task requires a MATTER_ID, but received CLIENT_ID: " + matterId, { page, matterId });
      }

      if (!matterId.startsWith("MAT_")) {
        return crm_renderErrorPage_("new_task requires a MATTER_ID; unknown id format: " + matterId, { page, matterId });
      }

      const matter = crm_getMatterById(matterId);
      if (!matter) return crm_renderErrorPage_("Matter not found: " + matterId, { page, matterId });

      const tpl = HtmlService.createTemplateFromFile("crm_app_new_task");
      tpl.page = page;
      tpl.matterId = matterId;
      tpl.clientId = matter.CLIENT_ID || "";
      tpl.matterTitle = matter.TITLE || matterId;
      tpl.clientName = tpl.clientId ? (crm_findClientById(tpl.clientId)?.FULL_NAME || "") : "";
      tpl.webAppBase = webAppBase;

      return tpl.evaluate()
        .setTitle("CRM New Task")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (page === "upload_doc") {
      const matterId = (e && e.parameter && e.parameter.matterId)
        ? String(e.parameter.matterId).trim()
        : "";

      if (!matterId) return crm_renderErrorPage_("Missing matterId for document upload", { page, matterId });

      const matter = crm_getMatterById(matterId);
      if (!matter) return crm_renderErrorPage_("Matter not found: " + matterId, { page, matterId });

      const tpl = HtmlService.createTemplateFromFile("crm_app_upload_doc");
      tpl.page = page;
      tpl.matterId = matterId;
      tpl.matterTitle = matter.TITLE || matterId;
      tpl.webAppBase = webAppBase;

      return tpl.evaluate()
        .setTitle("CRM Upload Document")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    const homeTpl = HtmlService.createTemplateFromFile("crm_app_shell");
    homeTpl.page = "home";
    homeTpl.entityId = id;
    homeTpl.webAppBase = webAppBase;

    return homeTpl.evaluate()
      .setTitle("Office CRM")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (err) {
    return crm_renderErrorPage_(err, {
      page: e && e.parameter ? e.parameter.page : "",
      id: e && e.parameter ? e.parameter.id : ""
    });
  }
}

function crm_renderErrorPage_(err, meta) {
  const msg = (err && err.message) ? err.message : String(err || "Unknown error");
  const details = JSON.stringify(meta || {}, null, 2);
  const webAppBase = ScriptApp.getService().getUrl();
  const homeUrl = webAppBase ? webAppBase + "?page=home" : "?page=home";

  return HtmlService.createHtmlOutput(`
    <!doctype html>
    <html>
      <head>
        <meta charset="utf-8" />
        <style>
          body { font-family: Arial, sans-serif; background:#f6f8fc; padding:24px; color:#1f2937; }
          .box { background:white; border:1px solid #e5e7eb; border-radius:16px; padding:20px; max-width:900px; margin:0 auto; }
          h1 { margin-top:0; color:#b91c1c; }
          pre { background:#f3f4f6; padding:12px; border-radius:12px; overflow:auto; }
          a { color:#1a73e8; }
        </style>
      </head>
      <body>
        <div class="box">
          <h1>CRM Web App Error</h1>
          <p><b>Message:</b> ${crm_escapeHtml_(msg)}</p>
          <p><b>Meta:</b></p>
          <pre>${crm_escapeHtml_(details)}</pre>
          <p><a href="${homeUrl}">Go Home</a></p>
        </div>
      </body>
    </html>
  `).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function crm_escapeHtml_(s) {
  return String(s ?? "").replace(/[&<>"']/g, function(c) {
    return {
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#39;"
    }[c];
  });
}

function crm_getWebAppUrl_(page, id) {
  const base = ScriptApp.getService().getUrl();
  if (!base) throw new Error("Web App URL unavailable. Deploy as Web App first.");

  let url = base + "?page=" + encodeURIComponent(page || "home");
  if (id) url += "&id=" + encodeURIComponent(id);
  return url;
}

function crm_webCreateLead(data) {
  const input = {
    fullName: data.FULL_NAME,
    fullNameRu: String(data.FULL_NAME_RU || data.FULL_NAME || "").trim(),
    fullNameHe: String(data.FULL_NAME_HE || "").trim(),
    phone: data.PHONE,
    source: data.SOURCE,
    caseType: data.CASE_TYPE,
    idType: String(data.ID_TYPE || "").trim(),
    idNumber: String(data.ID_NUMBER || "").trim(),
    address: String(data.ADDRESS || "").trim(),
    addressRu: String(data.ADDRESS_RU || data.ADDRESS || "").trim(),
    addressHe: String(data.ADDRESS_HE || "").trim(),
    subjectRu: String(data.SUBJECT_RU || "").trim(),
    subjectHe: String(data.SUBJECT_HE || "").trim(),
    eventDate: String(data.EVENT_DATE || "").trim(),
  };
  const result = crm_addLead(input);
  return { id: result.leadId };
}

// Returns lightweight client list for the New Matter dropdown.
function crm_webListClientsForSelect() {
  const c = cfg_();
  const ss = crm_getSpreadsheet_();
  const sh = ss.getSheetByName(c.SHEETS.CLIENTS);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const header = values[0];
  const iId = header.indexOf("CLIENT_ID");
  const iName = header.indexOf("FULL_NAME");
  const iPhone = header.indexOf("PHONE");
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const clientId = String(values[r][iId] || "").trim();
    if (!clientId) continue;
    const name = String(values[r][iName] || "").trim();
    const phone = String(values[r][iPhone] || "").trim();
    out.push({ clientId, label: name + (phone ? " (" + phone + ")" : "") });
  }
  return out;
}

function crm_webGetLeadInfo(leadId) {
  const lead = crm_getLeadById(leadId);
  if (!lead) return { error: "Lead not found", leadId };
  return {
    ok: true,
    lead: lead,
    hasLinkedClient: !!(lead.CLIENT_ID && String(lead.CLIENT_ID).trim()),
    linkedClientId: (lead.CLIENT_ID && String(lead.CLIENT_ID).trim()) || null,
  };
}

function crm_webCreateClient(data) {
  const input = {
    fullName: String(data.FULL_NAME ?? "").trim(),
    fullNameRu: String(data.FULL_NAME_RU || data.FULL_NAME || "").trim(),
    fullNameHe: String(data.FULL_NAME_HE ?? "").trim(),
    phone: String(data.PHONE ?? "").trim(),
    idNumber: String(data.ID_NUMBER ?? "").trim(),
    idType: String(data.ID_TYPE ?? "").trim(),
    address: String(data.ADDRESS ?? "").trim(),
    addressRu: String(data.ADDRESS_RU || data.ADDRESS || "").trim(),
    addressHe: String(data.ADDRESS_HE ?? "").trim(),
  };

  // If sourceLeadId provided, check if lead already has linked client
  if (data.sourceLeadId) {
    const lead = crm_getLeadById(data.sourceLeadId);
    if (lead && lead.CLIENT_ID && String(lead.CLIENT_ID).trim()) {
      return {
        id: lead.CLIENT_ID,
        alreadyLinked: true,
        message: "This lead is already linked to client: " + lead.CLIENT_ID,
        linkedClientId: lead.CLIENT_ID,
      };
    }
  }

  const result = crm_addClient(input);

  // If sourceLeadId provided, link the lead to this new client
  if (data.sourceLeadId && result.clientId) {
    const conv = crm_convertLeadToClient(data.sourceLeadId);
    // conversion already logged
  }

  return { id: result.clientId, alreadyLinked: false };
}

function crm_webUpdateClient(data) {
  if (!data || !data.clientId) throw new Error("crm_webUpdateClient: clientId is required");

  const input = {
    fullName: data.fullName,
    fullNameRu: data.fullNameRu,
    fullNameHe: data.fullNameHe,
    phone: data.phone,
    email: data.email,
    idNumber: data.idNumber,
    idType: data.idType,
    locale: data.locale,
    status: data.status,
    owner: data.owner,
    source: data.source,
    address: data.address,
    addressRu: data.addressRu,
    addressHe: data.addressHe,
  };

  const result = crm_updateClient(data.clientId, input);
  return { ok: true, clientId: result.CLIENT_ID };
}

function crm_webCreateMatterFromLead(leadId, data) {
  const lead = crm_getLeadById(leadId);
  if (!lead) throw new Error("Lead not found: " + leadId);

  // Resolve correct client_id: use existing linked client, or client_id from data
  let clientId = (lead.CLIENT_ID && String(lead.CLIENT_ID).trim()) || data.CLIENT_ID;
  if (!clientId) throw new Error("No client found: lead not linked and no CLIENT_ID provided");

  const input = {
    clientId: clientId,
    category: data.CATEGORY || lead.CASE_TYPE || "GENERAL",
    title: data.TITLE,
    authority: data.AUTHORITY || "",
    owner: data.OWNER || "",
  };

  const result = crm_createMatter(input);
  return { id: result.matterId };
}

function crm_webCreateMatter(data) {
  const input = {
    clientId: data.CLIENT_ID,
    category: data.CATEGORY,
    title: data.TITLE,
    authority: data.AUTHORITY,
    owner: data.OWNER,
    eventDate: String(data.EVENT_DATE || "").trim(),
    subjectRu: String(data.SUBJECT_RU || "").trim(),
    subjectHe: String(data.SUBJECT_HE || "").trim(),
  };
  const result = crm_createMatter(input);
  return { id: result.matterId };
}

function crm_webAddDocument(data) {
  // data: { clientId, matterId?, title, type?, docUrl?, pdfUrl?, notes? }
  const input = {
    clientId: data.clientId,
    matterId: data.matterId || "",
    title: data.title || "Document",
    type: data.type || "GENERAL",
    docUrl: data.docUrl || "",
    pdfUrl: data.pdfUrl || "",
    notes: data.notes || "",
  };
  const result = crm_addDocument(input);
  return { ok: result.ok, id: result.docId };
}

function crm_webCreateTask(data) {
  if (!data || !data.clientId || !data.matterId) {
    throw new Error("crm_webCreateTask: clientId and matterId are required");
  }

  const input = {
    clientId: data.clientId,
    matterId: data.matterId,
    title: data.title,
    type: data.type || "GENERAL",
    priority: data.priority || "MEDIUM",
    dueDate: data.dueDate || "",
    status: data.status || "OPEN",
    notes: data.notes || "",
    assignee: data.assignee || "",
    generatedBy: data.generatedBy || "MANUAL",
  };

  const result = crm_createTask(input);
  return { ok: result.ok, taskId: result.taskId };
}

function crm_webStartLeadOnboarding(leadId) {
  if (!leadId) throw new Error("crm_webStartLeadOnboarding: leadId is required");
  // Signing-first: creates onboarding matter with leadId as provisional CLIENT_ID.
  // No real client record is created here.
  // Lead STATUS → ONBOARDING now; real client created + STATUS → CONVERTED only after
  // successful signing via crm_finalizeOnboardingConversion_ inside crm_submitSignature.
  return crm_initLeadOnboardingContext_(leadId);
}

function crm_webUploadMatterDocument(data) {
  // data: { matterId, title, type, status, notes, fileName, mimeType, base64Data }
  const payload = {
    matterId: data.matterId,
    title: data.title,
    type: data.type || "GENERAL",
    status: data.status || "PENDING",
    notes: data.notes || "",
    fileName: data.fileName,
    mimeType: data.mimeType,
    base64Data: data.base64Data,
  };
  const result = crm_uploadFileToDriveAndRegister(payload);
  return {
    ok: result.ok,
    fileId: result.fileId,
    docId: result.docId,
    fileUrl: result.fileUrl,
    matterId: result.matterId,
    clientId: result.clientId,
  };
}