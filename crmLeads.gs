/** crmLeads.gs */

function crm_addLead(input) {
  const c = cfg_();
  const ss = crm_getSpreadsheet_();

  const sh = ss.getSheetByName(c.SHEETS.LEADS);
  if (!sh) throw new Error("Leads sheet missing");

  const fullName = (input?.fullName ?? "").trim();
  const phone = (input?.phone ?? "").trim();

  if (!fullName) throw new Error("crm_addLead: fullName is required");
  if (!phone) throw new Error("crm_addLead: phone is required");

  const leadId = generateId_("LEAD");
  const ts = nowIso_();

  const rowObj = {
    LEAD_ID: leadId,
    CREATED_AT: ts,
    UPDATED_AT: ts,
    SOURCE: (input?.source ?? "").trim(),
    CAMPAIGN: (input?.campaign ?? "").trim(),
    FULL_NAME: fullName,
    PHONE: phone,
    EMAIL: (input?.email ?? "").trim(),
    CASE_TYPE: (input?.caseType ?? "").trim(),
    DESCRIPTION: (input?.description ?? "").trim(),
    STATUS: (input?.status ?? "NEW").trim(),
    ASSIGNED_TO: (input?.assignedTo ?? getActiveUserEmail_()).trim(),
    CLIENT_ID: (input?.clientId ?? "").trim(),
    NOTES: (input?.notes ?? "").trim(),
    ID_TYPE: (input?.idType ?? "").trim(),
    ID_NUMBER: (input?.idNumber ?? "").trim(),
    ADDRESS: (input?.address ?? "").trim(),
  };

  const row = appendRowByHeaders_(sh, c.HEADERS.LEADS, rowObj);

  crm_logActivity({
    action: "LEAD_CREATED",
    message: `Lead created: ${fullName}`,
    meta: {
      leadId,
      phone: rowObj.PHONE,
      source: rowObj.SOURCE,
      caseType: rowObj.CASE_TYPE,
      row,
    },
  });

  logInfo_("LEAD", "Created lead", { leadId, row });

  return { ok: true, leadId, row };
}

function crm_getLeadById(leadId) {
  const ss = crm_getSpreadsheet_();
  const c = cfg_();

  const sh = ss.getSheetByName(c.SHEETS.LEADS);
  if (!sh) throw new Error("Leads sheet missing");
  if (!leadId) throw new Error("crm_getLeadById: leadId is required");

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return null;

  const header = values[0];
  const rows = values.slice(1);

  const iLead = header.indexOf("LEAD_ID");
  if (iLead === -1) throw new Error("LEAD_ID column not found");

  for (let r = 0; r < rows.length; r++) {
    if (String(rows[r][iLead]) === String(leadId)) {
      return rowToObj_(header, rows[r], r + 2);
    }
  }

  return null;
}

function crm_findLeadByPhone(phone) {
  const c = cfg_();
  const ss = crm_getSpreadsheet_();
  const sh = ss.getSheetByName(c.SHEETS.LEADS);

  if (!sh) throw new Error("Leads sheet not found");

  const target = normPhone_(phone);
  if (!target) return null;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

  const phoneCol = headers.indexOf("PHONE");
  if (phoneCol === -1) throw new Error("PHONE column not found in Leads headers");

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const p = normPhone_(row[phoneCol]);

    if (p && p === target) {
      return rowToObj_(headers, row, i + 2);
    }
  }

  return null;
}

function crm_updateLeadStatus(leadId, status) {
  if (!leadId) throw new Error("crm_updateLeadStatus: leadId is required");
  if (!status) throw new Error("crm_updateLeadStatus: status is required");

  const c = cfg_();
  const ss = crm_getSpreadsheet_();
  const sh = ss.getSheetByName(c.SHEETS.LEADS);
  if (!sh) throw new Error("Leads sheet missing");

  const values = sh.getDataRange().getValues();
  if (values.length < 2) throw new Error("crm_updateLeadStatus: no data in Leads");

  const header = values[0];
  const idx = (name) => header.indexOf(name);

  const iLeadId = idx("LEAD_ID");
  const iStatus = idx("STATUS");
  const iUpdatedAt = idx("UPDATED_AT");

  if (iLeadId < 0 || iStatus < 0 || iUpdatedAt < 0) {
    throw new Error("crm_updateLeadStatus: required columns missing");
  }

  const now = nowIso_();

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][iLeadId]) === String(leadId)) {
      const rowNum = r + 1;

      sh.getRange(rowNum, iStatus + 1).setValue(status);
      sh.getRange(rowNum, iUpdatedAt + 1).setValue(now);

      crm_logActivity({
        action: "LEAD_STATUS_UPDATED",
        message: `Lead status updated: ${leadId} -> ${status}`,
        meta: {
          leadId,
          status,
          row: rowNum,
        },
      });

      return { ok: true, leadId, status, row: rowNum };
    }
  }

  throw new Error("crm_updateLeadStatus: lead not found: " + leadId);
}

function crm_listLeads(filter) {
  filter = filter || {};

  const ss = crm_getSpreadsheet_();
  const c = cfg_();
  const sh = ss.getSheetByName(c.SHEETS.LEADS);
  if (!sh) throw new Error("crm_listLeads: Leads sheet not found");

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const header = values[0];
  const rows = values.slice(1);

  const idx = (name) => header.indexOf(name);

  const iStatus = idx("STATUS");
  const iSource = idx("SOURCE");
  const iAssignedTo = idx("ASSIGNED_TO");

  const statusNeed = filter.status ? String(filter.status) : "";
  const sourceNeed = filter.source ? String(filter.source) : "";
  const assignedNeed = filter.assignedTo ? String(filter.assignedTo) : "";
  const limit = Number(filter.limit || 50);

  const out = [];

  for (let r = rows.length - 1; r >= 0; r--) {
    const row = rows[r];

    const status = String(row[iStatus] || "");
    if (statusNeed && status !== statusNeed) continue;

    const source = String(row[iSource] || "");
    if (sourceNeed && source !== sourceNeed) continue;

    const assignedTo = String(row[iAssignedTo] || "");
    if (assignedNeed && assignedTo !== assignedNeed) continue;

    out.push(rowToObj_(header, row, r + 2));

    if (out.length >= limit) break;
  }

  return out;
}