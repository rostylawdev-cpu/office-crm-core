/** crmLeadService.gs */

function crm_upsertLead(input) {
  const phone = input?.phone ? String(input.phone) : "";
  if (!phone) throw new Error("crm_upsertLead: phone is required");

  const existing = crm_findLeadByPhone(phone);

  if (!existing) {
    return crm_addLead(input);
  }

  const c = cfg_();
  const ss = crm_getSpreadsheet_();
  const sh = ss.getSheetByName(c.SHEETS.LEADS);
  if (!sh) throw new Error("Leads sheet missing");

  const row = existing.__row;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = (name) => headers.indexOf(name) + 1;

  const patch = {
    SOURCE: input.source,
    CAMPAIGN: input.campaign,
    FULL_NAME: input.fullName,
    EMAIL: input.email,
    CASE_TYPE: input.caseType,
    DESCRIPTION: input.description,
    STATUS: input.status,
    ASSIGNED_TO: input.assignedTo,
    NOTES: input.notes,
    UPDATED_AT: nowIso_(),
  };

  let changed = false;

  Object.keys(patch).forEach((k) => {
    const v = patch[k];
    if (v === undefined || v === null || v === "") return;

    const colIdx = col(k);
    if (colIdx <= 0) return;

    const current = sh.getRange(row, colIdx).getValue();

    if (current === "" || current === null) {
      sh.getRange(row, colIdx).setValue(v);
      existing[k] = v;
      changed = true;
    } else if (k === "UPDATED_AT") {
      sh.getRange(row, colIdx).setValue(v);
      existing[k] = v;
      changed = true;
    }
  });

  if (changed) {
    crm_logActivity({
      action: "LEAD_UPDATED",
      message: `Lead updated: ${existing.LEAD_ID}`,
      meta: {
        leadId: existing.LEAD_ID,
        phone: normPhone_(phone),
        row,
      },
    });
  }

  return existing;
}

function crm_convertLeadToClient(leadId) {
  const lead = crm_getLeadById(leadId);
  if (!lead) throw new Error("crm_convertLeadToClient: lead not found: " + leadId);

  if (lead.CLIENT_ID) {
    return {
      ok: true,
      leadId,
      clientId: lead.CLIENT_ID,
      alreadyLinked: true,
    };
  }

  const clientRes = crm_upsertClient({
    fullName: lead.FULL_NAME,
    phone: lead.PHONE,
    email: lead.EMAIL,
    source: lead.SOURCE || "LEAD_CONVERSION",
    status: "NEW",
    owner: lead.ASSIGNED_TO || getActiveUserEmail_(),
  });

  const clientId = clientRes.clientId || clientRes.CLIENT_ID;
  if (!clientId) {
    throw new Error("crm_convertLeadToClient: failed to get clientId after conversion");
  }

  const c = cfg_();
  const ss = crm_getSpreadsheet_();
  const sh = ss.getSheetByName(c.SHEETS.LEADS);
  if (!sh) throw new Error("Leads sheet missing");

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = (name) => headers.indexOf(name) + 1;

  const row = lead.__row;
  const now = nowIso_();

  const iClient = col("CLIENT_ID");
  const iStatus = col("STATUS");
  const iUpdated = col("UPDATED_AT");

  if (iClient > 0) sh.getRange(row, iClient).setValue(clientId);
  if (iStatus > 0) sh.getRange(row, iStatus).setValue("CONVERTED");
  if (iUpdated > 0) sh.getRange(row, iUpdated).setValue(now);

  crm_logActivity({
    action: "LEAD_CONVERTED_TO_CLIENT",
    message: `Lead converted to client: ${leadId} -> ${clientId}`,
    clientId: clientId,
    meta: {
      leadId,
      clientId,
      leadRow: row,
    },
  });

  return {
    ok: true,
    leadId,
    clientId,
    row,
  };
}

function crm_convertLeadToClientAndMatter(leadId, matterInput) {
  const lead = crm_getLeadById(leadId);
  if (!lead) throw new Error("crm_convertLeadToClientAndMatter: lead not found: " + leadId);

  const conv = crm_convertLeadToClient(leadId);
  const clientId = conv.clientId;

  const matterRes = crm_createMatterWithTasks({
    clientId: clientId,
    category: matterInput?.category || lead.CASE_TYPE || "GENERAL",
    title: matterInput?.title || `Matter from lead ${lead.FULL_NAME}`,
    authority: matterInput?.authority || "",
    stage: matterInput?.stage || "NEW",
    taskTemplateKey: matterInput?.taskTemplateKey || "",
    owner: matterInput?.owner || lead.ASSIGNED_TO || "",
  });

  crm_logActivity({
    action: "LEAD_CONVERTED_TO_MATTER",
    message: `Lead converted to client + matter: ${leadId}`,
    clientId: clientId,
    matterId: matterRes.matterId,
    meta: {
      leadId,
      clientId,
      matterId: matterRes.matterId,
    },
  });

  return {
    ok: true,
    leadId,
    clientId,
    matterId: matterRes.matterId,
    tasksCreated: matterRes.tasksCreated || 0,
  };
}