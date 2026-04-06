/** crmClients.gs */

/**
 * Create (C in CRUD)
 * @param {Object} input
 * @param {string} input.fullName
 * @param {string} input.phone
 * @param {string} [input.email]
 * @param {string} [input.idNumber]
 * @param {string} [input.locale]
 * @param {string} [input.status]
 * @param {string} [input.owner]
 * @param {string} [input.source]
 * @returns {{clientId:string, row:number}}
 */
function crm_addClient(input) {
  const c = cfg_();
  const ss = crm_getSpreadsheet_();

  const sh = ss.getSheetByName(c.SHEETS.CLIENTS);
  if (!sh) throw new Error("Clients sheet missing");

  const fullName = String(input?.fullName ?? "").trim();
  const phone = String(input?.phone ?? "").trim();

  if (!fullName) throw new Error("fullName is required");
  if (!phone) throw new Error("phone is required");

  // Duplicate guard: return existing client if name + phone match
  const duplicate = crm_findClientByIdentity_(fullName, phone);
  if (duplicate) {
    logInfo_("CLIENT_DEDUP", "crm_addClient: returning existing client (name+phone match)", {
      clientId: duplicate.CLIENT_ID, fullName, phone
    });
    return { clientId: duplicate.CLIENT_ID, row: duplicate.__row, existing: true };
  }

  const clientId = generateId_("CL");
  const ts = nowIso_();

  // Create client folder (safe: returns null if folder creation fails)
  const folderUrl = crm_getOrCreateClientFolder(clientId, fullName);

  const rowObj = {
    CLIENT_ID: clientId,
    FULL_NAME: fullName,
    PHONE: phone,
    EMAIL: String(input?.email ?? "").trim(),
    ID_NUMBER: String(input?.idNumber ?? "").trim(),
    ID_TYPE: crm_normalizeIdType_(input?.idType ?? ""),
    LOCALE: String(input?.locale ?? "").trim(),
    STATUS: String(input?.status ?? "NEW").trim(),
    OWNER: String(input?.owner ?? getActiveUserEmail_()).trim(),
    SOURCE: String(input?.source ?? "").trim(),
    ADDRESS: String(input?.address ?? "").trim(),
    FULL_NAME_RU: String(input?.fullNameRu ?? fullName).trim(),
    FULL_NAME_HE: String(input?.fullNameHe ?? "").trim(),
    ADDRESS_RU: String(input?.addressRu ?? input?.address ?? "").trim(),
    ADDRESS_HE: String(input?.addressHe ?? "").trim(),
    FOLDER_URL: folderUrl || "",
    CREATED_AT: ts,
    UPDATED_AT: ts,
    LAST_ACTIVITY_AT: ts,
  };

  const row = appendRowByHeaders_(sh, c.HEADERS.CLIENTS, rowObj);

  crm_logActivity({
    clientId,
    action: "CLIENT_CREATED",
    message: `Client created: ${fullName}`,
    meta: { phone, email: rowObj.EMAIL, row, folderUrl: folderUrl || "" },
  });

  logInfo_("CLIENT", "Created client", { clientId, row, folderUrl: folderUrl || "" });

  return { clientId, row };
}

/**
 * Finds a client by full name + phone (both normalized).
 * Used to prevent duplicate client creation when the same person signs up twice.
 */
function crm_findClientByIdentity_(fullName, phone) {
  const rows = crm_getAllRowsFromSheet_(cfg_().SHEETS.CLIENTS, cfg_().HEADERS.CLIENTS);
  const targetName = String(fullName || "").trim().toLowerCase();
  const targetPhone = normPhone_(phone);
  return rows.find(function(r) {
    return String(r.FULL_NAME || "").trim().toLowerCase() === targetName &&
           normPhone_(r.PHONE) === targetPhone;
  }) || null;
}

function crm_findClientByPhone(phone) {
  const c = cfg_();
  const ss = crm_getSpreadsheet_();
  const sh = ss.getSheetByName(c.SHEETS.CLIENTS);

  if (!sh) throw new Error("Clients sheet not found");

  const target = normPhone_(phone);
  if (!target) return null;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

  const phoneCol = headers.indexOf("PHONE");
  if (phoneCol === -1) throw new Error("PHONE column not found in Clients headers");

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const p = normPhone_(row[phoneCol]);

    if (p && p === target) {
      return rowToObj_(headers, row, i + 2);
    }
  }

  return null;
}

function crm_findClientById(clientId) {
  const ss = crm_getSpreadsheet_();
  const c = cfg_();

  const sh = ss.getSheetByName(c.SHEETS.CLIENTS);
  if (!sh) throw new Error("Clients sheet missing");
  if (!clientId) throw new Error("crm_findClientById: clientId is required");

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return null;

  const header = values[0];
  const rows = values.slice(1);

  const iClient = header.indexOf("CLIENT_ID");
  if (iClient === -1) throw new Error("CLIENT_ID column not found");

  for (let r = 0; r < rows.length; r++) {
    if (String(rows[r][iClient]) === String(clientId)) {
      return rowToObj_(header, rows[r], r + 2);
    }
  }

  return null;
}

function crm_upsertClient(p) {
  const phone = p && p.phone ? p.phone : "";
  if (!phone) throw new Error("crm_upsertClient: phone is required");

  const existing = crm_findClientByPhone(phone);

  if (!existing) {
    return crm_addClient(p);
  }

  const c = cfg_();
  const ss = crm_getSpreadsheet_();
  const sh = ss.getSheetByName(c.SHEETS.CLIENTS);
  if (!sh) throw new Error("Clients sheet missing");

  const row = existing.__row;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = (name) => headers.indexOf(name) + 1;

  const patch = {
    FULL_NAME: p.fullName !== undefined && p.fullName !== null ? String(p.fullName).trim() : "",
    EMAIL: p.email !== undefined && p.email !== null ? String(p.email).trim() : "",
    ID_TYPE: p.idType !== undefined && p.idType !== null ? crm_normalizeIdType_(p.idType) : "",
    ID_NUMBER: p.idNumber !== undefined && p.idNumber !== null ? String(p.idNumber).trim() : "",
    LOCALE: p.locale !== undefined && p.locale !== null ? String(p.locale).trim() : "",
    SOURCE: p.source !== undefined && p.source !== null ? String(p.source).trim() : "",
    OWNER: p.owner !== undefined && p.owner !== null ? String(p.owner).trim() : "",
    STATUS: p.status !== undefined && p.status !== null ? String(p.status).trim() : "",
    UPDATED_AT: nowIso_(),
    LAST_ACTIVITY_AT: nowIso_(),
  };

  let changed = false;

  Object.keys(patch).forEach((k) => {
    const v = patch[k];
    if (v === undefined || v === null || v === "") return;

    const colIdx = col(k);
    if (colIdx <= 0) return;

    const current = sh.getRange(row, colIdx).getValue();

    // не перезаписываем непустое, кроме служебных timestamp-полей
    if (current === "" || current === null) {
      sh.getRange(row, colIdx).setValue(v);
      existing[k] = v;
      changed = true;
    } else if (k === "UPDATED_AT" || k === "LAST_ACTIVITY_AT") {
      sh.getRange(row, colIdx).setValue(v);
      existing[k] = v;
      changed = true;
    }
  });

  if (changed) {
    logInfo_("CLIENT_UPDATE", "Client updated (upsert)", {
      clientId: existing.CLIENT_ID,
      phone: normPhone_(phone),
      row,
    });
  } else {
    logInfo_("CLIENT_EXISTS", "Client exists, nothing to update", {
      clientId: existing.CLIENT_ID,
      phone: normPhone_(phone),
      row,
    });
  }

  return existing;
}

function crm_updateClient(clientId, updates) {
  if (!clientId) throw new Error("crm_updateClient: clientId is required");
  if (!updates || typeof updates !== "object") throw new Error("crm_updateClient: updates object is required");

  const existing = crm_findClientById(clientId);
  if (!existing) throw new Error("crm_updateClient: client not found: " + clientId);

  const ss = crm_getSpreadsheet_();
  const c = cfg_();
  const sh = ss.getSheetByName(c.SHEETS.CLIENTS);
  if (!sh) throw new Error("crm_updateClient: Clients sheet not found");

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const rowNum = existing.__row;
  if (!rowNum) throw new Error("crm_updateClient: existing row not found");

  const fieldValues = {
    FULL_NAME: updates.fullName,
    FULL_NAME_RU: updates.fullNameRu,
    FULL_NAME_HE: updates.fullNameHe,
    PHONE: updates.phone,
    EMAIL: updates.email,
    ID_NUMBER: updates.idNumber,
    ID_TYPE: updates.idType,
    LOCALE: updates.locale,
    STATUS: updates.status,
    OWNER: updates.owner,
    SOURCE: updates.source,
    ADDRESS: updates.address,
    ADDRESS_RU: updates.addressRu,
    ADDRESS_HE: updates.addressHe,
  };

  let changed = false;

  Object.keys(fieldValues).forEach(function (field) {
    const value = fieldValues[field];
    if (value === undefined || value === null) return;
    const trimmed = String(value).trim();
    if (trimmed === "") return;

    const colIdx = headers.indexOf(field) + 1;
    if (colIdx <= 0) return;

    sh.getRange(rowNum, colIdx).setValue(trimmed);
    existing[field] = trimmed;
    changed = true;
  });

  const now = nowIso_();
  const updatedAtCol = headers.indexOf("UPDATED_AT") + 1;
  if (updatedAtCol > 0) {
    sh.getRange(rowNum, updatedAtCol).setValue(now);
    existing.UPDATED_AT = now;
    changed = true;
  }
  const lastActivityCol = headers.indexOf("LAST_ACTIVITY_AT") + 1;
  if (lastActivityCol > 0) {
    sh.getRange(rowNum, lastActivityCol).setValue(now);
    existing.LAST_ACTIVITY_AT = now;
    changed = true;
  }

  if (changed) {
    crm_logActivity({
      action: "CLIENT_UPDATED",
      message: "Client updated: " + clientId,
      clientId: clientId,
      meta: { updates: fieldValues },
    });
  }

  return existing;
}
