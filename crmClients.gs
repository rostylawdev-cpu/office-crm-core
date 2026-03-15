/** crmClients.js */

function crm_getSpreadsheet_() {
  const c = cfg_();
  const ssId = getProp_(c.CRM_SPREADSHEET_ID_PROP);
  if (!ssId) throw new Error(`Missing Script Property: ${c.CRM_SPREADSHEET_ID_PROP}`);
  return SpreadsheetApp.openById(ssId);
}

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

  const fullName = (input?.fullName ?? "").trim();
  const phone = (input?.phone ?? "").trim();

  if (!fullName) throw new Error("fullName is required");
  if (!phone) throw new Error("phone is required");

  const clientId = generateId_("CL");
  const ts = nowIso_();

  const rowObj = {
    CLIENT_ID: clientId,
    FULL_NAME: fullName,
    PHONE: phone,
    EMAIL: (input?.email ?? "").trim(),
    ID_NUMBER: (input?.idNumber ?? "").trim(),
    LOCALE: (input?.locale ?? "").trim(),
    STATUS: (input?.status ?? "NEW").trim(),
    OWNER: (input?.owner ?? getActiveUserEmail_()).trim(),
    SOURCE: (input?.source ?? "").trim(),
    CREATED_AT: ts,
    UPDATED_AT: ts,
    LAST_ACTIVITY_AT: ts,
    // FOLDER_URL / PAYMENT_* пока пустые
  };

  const row = appendRowByHeaders_(sh, c.HEADERS.CLIENTS, rowObj);

  crm_logActivity_(ss, {
    clientId,
    action: "CLIENT_CREATED",
    message: `Client created: ${fullName}`,
    meta: { phone, email: rowObj.EMAIL, row },
  });

  logInfo_("CLIENT", "Created client", { clientId, row });

  return { clientId, row };
}

function crm_upsertClient(p) {
  const phone = p && p.phone ? p.phone : "";
  if (!phone) throw new Error("crm_upsertClient: phone is required");

  const existing = crm_findClientByPhone(phone);

  if (!existing) {
    // Create
    return crm_addClient(p); // твоя уже рабочая функция
  }

  // Update (только если в существующем пусто)
  const c = cfg_();
  const ssId = getProp_(c.CRM_SPREADSHEET_ID_PROP);
  if (!ssId) throw new Error("Missing Script Property: CRM_SPREADSHEET_ID");
  const ss = SpreadsheetApp.openById(ssId);
  const sh = ss.getSheetByName(c.SHEETS.CLIENTS);

  const row = existing.__row;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = (name) => headers.indexOf(name) + 1;

  const patch = {
    FULL_NAME: p.fullName,
    EMAIL: p.email,
    LOCALE: p.locale,
    SOURCE: p.source,
    OWNER: p.owner,
    STATUS: p.status,
    UPDATED_AT: nowIso_(),
    LAST_ACTIVITY_AT: nowIso_(),
  };

  let changed = false;

  Object.keys(patch).forEach((k) => {
    const v = patch[k];
    if (v === undefined || v === null || v === "") return;

    const colIdx = col(k);
    if (colIdx <= 0) return;

    // правило "не перезаписываем непустое"
    const current = sh.getRange(row, colIdx).getValue();
    if (current === "" || current === null) {
      sh.getRange(row, colIdx).setValue(v);
      changed = true;
      existing[k] = v;
    } else if (k === "UPDATED_AT" || k === "LAST_ACTIVITY_AT") {
      // эти поля можно всегда обновлять
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

function test_upsertClient() {
  return crm_upsertClient({
    phone: "050-1234567",
    email: "darwin+new@example.com",
    fullName: "Darwin Saindrom",
    source: "TEST_UPSERT",
  });
}