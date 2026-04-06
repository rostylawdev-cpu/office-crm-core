/** Utils.gs */

// хелпер перенесен из crmClients
function crm_getSpreadsheet_() {
  const c = cfg_();
  const ssId = getProp_(c.CRM_SPREADSHEET_ID_PROP);
  if (!ssId) throw new Error(`Missing Script Property: ${c.CRM_SPREADSHEET_ID_PROP}`);
  return SpreadsheetApp.openById(ssId);
}

/**
 * Normalizes free-text ID type values to canonical CRM values: "TZ" or "PASSPORT".
 * Returns "" for unrecognized input.
 */
function crm_normalizeIdType_(v) {
  const s = String(v || "").toUpperCase().replace(/[\s_-]+/g, "");
  if (s === "TZ" || s === "TEUDATZEHUT") return "TZ";
  if (s === "PASSPORT" || s === "DARKON") return "PASSPORT";
  return "";
}

function getProp_(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

function setProp_(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, String(value));
}

function nowIso_() {
  return new Date().toISOString();
}

function json_(obj) {
  return JSON.stringify(obj ?? {});
}

function generateId_(prefix) {
  // Простой, читаемый ID: CL_20260304_114512_8F3K2A
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  const ts =
    d.getFullYear() +
    pad(d.getMonth() + 1) +
    pad(d.getDate()) + "_" +
    pad(d.getHours()) +
    pad(d.getMinutes()) +
    pad(d.getSeconds());

  const rnd = Math.random().toString(36).slice(2, 8).toUpperCase();
  return `${prefix}_${ts}_${rnd}`;
}

function getActiveUserEmail_() {
  // В consumer Gmail иногда возвращает пусто — это нормально
  try {
    return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || "";
  } catch (e) {
    return "";
  }
}

function ensureSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function ensureHeaders_(sh, headers) {
  if (!sh) throw new Error("ensureHeaders_: sheet is null");

  const existing =
    sh.getLastRow() >= 1
      ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
      : [];

  const needInit = sh.getLastRow() === 0 || existing.filter(Boolean).length === 0;

  if (needInit) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    return;
  }

  // Append any new headers not yet present in the sheet (append-only, never reorders existing columns)
  const existingSet = new Set(existing.map(function(h) { return String(h); }));
  const toAdd = headers.filter(function(h) { return h && !existingSet.has(h); });
  if (toAdd.length > 0) {
    const startCol = sh.getLastColumn() + 1;
    sh.getRange(1, startCol, 1, toAdd.length).setValues([toAdd]);
  }
}

function buildHeaderIndex_(headers) {
  const idx = {};
  headers.forEach((h, i) => {
    idx[h] = i;
  });
  return idx;
}

function appendRowByHeaders_(sheet, headers, obj) {
  const values = headers.map((h) => (obj[h] !== undefined ? obj[h] : ""));
  const row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, values.length).setValues([values]);
  return row;
}

function rowToObj_(headers, rowVals, rowNumber) {
  const o = { __row: rowNumber };
  headers.forEach((h, i) => {
    o[h] = rowVals[i];
  });
  return o;
}

function logInfo_(tag, message, meta) {
  console.log(`${tag}: ${message} ${json_(meta)}`);
}

function normPhone_(phone) {
  if (!phone) return "";

  let p = String(phone).trim();

  // убираем всё кроме цифр и +
  p = p.replace(/[^\d+]/g, "");

  // 00xx -> +xx
  if (p.startsWith("00")) p = "+" + p.slice(2);

  // +9725XXXXXXXX -> 05XXXXXXXX
  if (p.startsWith("+972")) {
    let rest = p.slice(4);
    if (rest.startsWith("0")) rest = rest.slice(1);
    p = "0" + rest;
  }

  // 9725XXXXXXXX -> 05XXXXXXXX
  if (p.startsWith("972")) {
    let rest = p.slice(3);
    if (rest.startsWith("0")) rest = rest.slice(1);
    p = "0" + rest;
  }

  return p;
}

function tryTouchClientLastActivity_(clientId, nowIso) {
  const ss = crm_getSpreadsheet_();
  const c = cfg_();
  const sh = ss.getSheetByName(c.SHEETS.CLIENTS);
  if (!sh) return;

  const headers = c.HEADERS.CLIENTS;
  const idxClientId = headers.indexOf("CLIENT_ID");
  const idxLast = headers.indexOf("LAST_ACTIVITY_AT");
  const idxUpdated = headers.indexOf("UPDATED_AT");

  if (idxClientId === -1 || idxLast === -1) return;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const data = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][idxClientId]) === String(clientId)) {
      const rowNumber = i + 2;
      sh.getRange(rowNumber, idxLast + 1).setValue(nowIso);
      if (idxUpdated !== -1) {
        sh.getRange(rowNumber, idxUpdated + 1).setValue(nowIso);
      }
      return;
    }
  }
}

function addDaysIso_(days) {
  const d = new Date();
  d.setDate(d.getDate() + Number(days || 0));
  return d.toISOString();
}

function escapeHtml_(s) {
  s = String(s ?? "");
  return s.replace(/[&<>"']/g, (c) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    "\"": "&quot;",
    "'": "&#39;"
  }[c]));
}

function getRowValueByHeader_(sheet, rowNumber, headerName) {
  if (!sheet) throw new Error("getRowValueByHeader_: sheet is required");
  if (!rowNumber || rowNumber < 1) throw new Error("getRowValueByHeader_: invalid rowNumber");
  if (!headerName) throw new Error("getRowValueByHeader_: headerName is required");

  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return "";

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = headers.indexOf(headerName);
  if (idx === -1) return "";

  return sheet.getRange(rowNumber, idx + 1).getValue();
}

function crm_getAllRowsFromSheet_(sheetName, headers) {
  const ss = crm_getSpreadsheet_();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) {
    logInfo_("DATA", "crm_getAllRowsFromSheet_: sheet not found: " + sheetName, {});
    return [];
  }

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const header = values[0];
  const rows = values.slice(1);

  return rows
    .filter(function (row) {
      return row.some(function (cell) { return cell !== ""; });
    })
    .map(function (row, i) {
      return rowToObj_(header, row, i + 2);
    });
}

/**
 * Extract folder ID from Google Drive URL.
 * Safe: returns null if URL is invalid or not a Drive folder URL.
 */
function extractFolderIdFromUrl_(folderUrl) {
  if (!folderUrl) return null;

  try {
    const url = String(folderUrl).trim();

    // Format: https://drive.google.com/drive/folders/{FOLDER_ID}
    const match = url.match(/\/folders\/([a-zA-Z0-9_-]+)/);
    if (match && match[1]) {
      return match[1];
    }

    // Format: https://drive.google.com/open?id={FOLDER_ID}
    const match2 = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (match2 && match2[1]) {
      return match2[1];
    }

    return null;
  } catch (e) {
    return null;
  }
}

/**
 * Extract ID from Google Drive/Docs URL (generic for both folders and files).
 * Handles:
 *   - /d/{ID} format (Google Docs)
 *   - ?id= or &id= query param format
 *   - /folders/{ID} format (Google Drive folders)
 */
function extractIdFromUrl_(url) {
  if (!url) return null;

  try {
    const str = String(url).trim();

    // Format: https://docs.google.com/document/d/FILE_ID/edit
    const match1 = str.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (match1 && match1[1]) {
      return match1[1];
    }

    // Format: ?id=FILE_ID or &id=FILE_ID
    const match2 = str.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (match2 && match2[1]) {
      return match2[1];
    }

    // Format: /folders/FOLDER_ID
    const match3 = str.match(/\/folders\/([a-zA-Z0-9_-]+)/);
    if (match3 && match3[1]) {
      return match3[1];
    }

    return null;
  } catch (e) {
    return null;
  }
}