/** crmDocuments.gs */

function crm_addDocument(input) {
  // input: {
  //   clientId, matterId?, type?, status?, title,
  //   docUrl?, pdfUrl?, fileId?, createdBy?, notes?
  // }

  const ss = crm_getSpreadsheet_();
  const c = cfg_();

  if (!input || !input.clientId) {
    throw new Error("crm_addDocument: missing input.clientId");
  }
  if (!input.title) {
    throw new Error("crm_addDocument: missing input.title");
  }

  const sh = ss.getSheetByName(c.SHEETS.DOCUMENTS);
  if (!sh) throw new Error("crm_addDocument: DOCUMENTS sheet not found");

  const now = nowIso_();
  const actor = input.createdBy || getActiveUserEmail_() || "unknown";
  const docId = generateId_("DOC");

  const rowObj = {
    DOC_ID: docId,
    CLIENT_ID: input.clientId,
    MATTER_ID: input.matterId || "",
    TYPE: (input.type || "GENERAL").toUpperCase(),
    STATUS: (input.status || "DRAFT").toUpperCase(),
    TITLE: input.title,
    DOC_URL: input.docUrl || "",
    PDF_URL: input.pdfUrl || "",
    FILE_ID: input.fileId || "",
    CREATED_AT: now,
    CREATED_BY: actor,
    NOTES: input.notes || "",
  };

  const rowIndex = appendRowByHeaders_(sh, c.HEADERS.DOCUMENTS, rowObj);

  crm_logActivity({
    action: "DOCUMENT_CREATED",
    message: `Document created: ${docId}`,
    clientId: input.clientId,
    matterId: input.matterId || "",
    meta: {
      row: rowIndex,
      docId: docId,
      type: rowObj.TYPE,
      status: rowObj.STATUS,
      title: rowObj.TITLE,
    },
  });

  tryTouchClientLastActivity_(input.clientId, now);

  return {
    ok: true,
    docId,
    row: rowIndex,
  };
}

function crm_getDocumentById(docId) {
  const ss = crm_getSpreadsheet_();
  const c = cfg_();

  const sh = ss.getSheetByName(c.SHEETS.DOCUMENTS);
  if (!sh) throw new Error("crm_getDocumentById: DOCUMENTS sheet not found");
  if (!docId) throw new Error("crm_getDocumentById: missing docId");

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return null;

  const header = values[0];
  const rows = values.slice(1);

  const iDoc = header.indexOf("DOC_ID");
  if (iDoc === -1) throw new Error("crm_getDocumentById: DOC_ID column not found");

  for (let r = 0; r < rows.length; r++) {
    if (String(rows[r][iDoc]) === String(docId)) {
      return rowToObj_(header, rows[r], r + 2);
    }
  }

  return null;
}

function crm_listDocumentsByClientId(clientId, opt) {
  opt = opt || {};
  const limit = Number(opt.limit || 50);

  if (!clientId) throw new Error("crm_listDocumentsByClientId: missing clientId");

  const rows = crm_getAllRowsFromSheet_(cfg_().SHEETS.DOCUMENTS, cfg_().HEADERS.DOCUMENTS) || [];

  const out = rows.filter(function (r) {
    return String(r.CLIENT_ID || "") === String(clientId);
  });

  out.sort(function (a, b) {
    const ta = a.CREATED_AT ? new Date(a.CREATED_AT).getTime() : 0;
    const tb = b.CREATED_AT ? new Date(b.CREATED_AT).getTime() : 0;
    return tb - ta;
  });

  return out.slice(0, limit);
}

function crm_listDocumentsByMatterId(matterId, opt) {
  opt = opt || {};
  const limit = Number(opt.limit || 50);

  if (!matterId) throw new Error("crm_listDocumentsByMatterId: missing matterId");

  const rows = crm_getAllRowsFromSheet_(cfg_().SHEETS.DOCUMENTS, cfg_().HEADERS.DOCUMENTS) || [];

  const out = rows.filter(function (r) {
    return String(r.MATTER_ID || "") === String(matterId);
  });

  out.sort(function (a, b) {
    const ta = a.CREATED_AT ? new Date(a.CREATED_AT).getTime() : 0;
    const tb = b.CREATED_AT ? new Date(b.CREATED_AT).getTime() : 0;
    return tb - ta;
  });

  return out.slice(0, limit);
}

function crm_updateDocumentStatus(docId, status, note) {
  if (!docId) throw new Error("crm_updateDocumentStatus: missing docId");
  if (!status) throw new Error("crm_updateDocumentStatus: missing status");

  const ss = crm_getSpreadsheet_();
  const c = cfg_();
  const sh = ss.getSheetByName(c.SHEETS.DOCUMENTS);
  if (!sh) throw new Error("crm_updateDocumentStatus: DOCUMENTS sheet not found");

  const values = sh.getDataRange().getValues();
  if (values.length < 2) throw new Error("crm_updateDocumentStatus: no data in DOCUMENTS");

  const header = values[0];
  const idx = function (name) { return header.indexOf(name); };

  const iDocId = idx("DOC_ID");
  const iStatus = idx("STATUS");
  const iNotes = idx("NOTES");
  const iClientId = idx("CLIENT_ID");
  const iMatterId = idx("MATTER_ID");
  const iTitle = idx("TITLE");

  if (iDocId < 0 || iStatus < 0) {
    throw new Error("crm_updateDocumentStatus: required columns missing");
  }

  let rowNum = -1;
  let row = null;

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][iDocId] || "") === String(docId)) {
      rowNum = r + 1;
      row = values[r];
      break;
    }
  }

  if (rowNum < 0) throw new Error("crm_updateDocumentStatus: document not found: " + docId);

  sh.getRange(rowNum, iStatus + 1).setValue(String(status).toUpperCase());

  if (iNotes >= 0 && note) {
    const prev = String(row[iNotes] || "");
    const next = prev ? prev + "\n" + note : String(note);
    sh.getRange(rowNum, iNotes + 1).setValue(next);
  }

  const clientId = iClientId >= 0 ? String(row[iClientId] || "") : "";
  const matterId = iMatterId >= 0 ? String(row[iMatterId] || "") : "";
  const title = iTitle >= 0 ? String(row[iTitle] || "") : "";

  crm_logActivity({
    action: "DOCUMENT_STATUS_UPDATED",
    message: `Document status updated: ${docId} -> ${String(status).toUpperCase()}`,
    clientId: clientId,
    matterId: matterId,
    meta: {
      row: rowNum,
      docId: docId,
      title: title,
      status: String(status).toUpperCase(),
      note: note || "",
    },
  });

  if (clientId) {
    tryTouchClientLastActivity_(clientId, nowIso_());
  }

  return {
    ok: true,
    docId,
    row: rowNum,
    status: String(status).toUpperCase(),
  };
}