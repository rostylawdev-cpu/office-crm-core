/** Utils.js */

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
  const existing = sh.getLastRow() >= 1 ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
  const needInit = sh.getLastRow() === 0 || existing.filter(Boolean).length === 0;

  if (needInit) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    return;
  }

  // Если заголовки уже есть — не трогаем (бережём руками правленные таблицы)
}

function buildHeaderIndex_(headers) {
  const idx = {};
  headers.forEach((h, i) => (idx[h] = i));
  return idx;
}

function appendRowByHeaders_(sh, headers, obj) {
  const idx = buildHeaderIndex_(headers);
  const row = new Array(headers.length).fill("");

  Object.keys(obj).forEach((k) => {
    if (k in idx) row[idx[k]] = obj[k] ?? "";
  });

  sh.appendRow(row);
  return sh.getLastRow();
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
    let rest = p.slice(4); // после +972
    if (rest.startsWith("0")) rest = rest.slice(1);
    p = "0" + rest; // обычно 0 + 5...
  }

  // 9725XXXXXXXX -> 05XXXXXXXX
  if (p.startsWith("972")) {
    let rest = p.slice(3);
    if (rest.startsWith("0")) rest = rest.slice(1);
    p = "0" + rest;
  }

  return p;
}

function crm_findClientByPhone(phone) {
  const c = cfg_();
  const ssId = getProp_(c.CRM_SPREADSHEET_ID_PROP);
  if (!ssId) throw new Error("Missing Script Property: CRM_SPREADSHEET_ID");

  const ss = SpreadsheetApp.openById(ssId);
  const sh = ss.getSheetByName(c.SHEETS.CLIENTS);
  if (!sh) throw new Error("Clients sheet not found");

  const target = normPhone_(phone);
  if (!target) return null;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null; // только заголовки

  // читаем все данные одним куском (быстро)
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

  const phoneCol = headers.indexOf("PHONE");
  if (phoneCol === -1) throw new Error("PHONE column not found in Clients headers");

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const p = normPhone_(row[phoneCol]);
    if (p && p === target) {
      const obj = {};
      headers.forEach((h, idx) => (obj[h] = row[idx]));
      obj.__row = i + 2; // реальный номер строки в Sheet
      return obj;
    }
  }

  return null;
}

function test_findClient() {
  const found = crm_findClientByPhone("050-1234567");
  Logger.log(found);
  return found ? found.CLIENT_ID : "NOT_FOUND";
}

function makeId_(prefix) {
  const d = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
  const rnd = Utilities.getUuid().slice(0, 6).toUpperCase();
  return `${prefix}_${d}_${rnd}`;
}

function appendRowByHeaders_(sheet, headers, obj) {
  const values = headers.map(h => (obj[h] !== undefined ? obj[h] : ""));
  const row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, values.length).setValues([values]);
  return row;
}

function rowToObj_(headers, rowVals, rowNumber) {
  const o = { __row: rowNumber };
  headers.forEach((h, i) => (o[h] = rowVals[i]));
  return o;
}

function logActivity_(p) {
  const ss = crm_getSpreadsheet_();
  const c = cfg_();
  const sh = ss.getSheetByName(c.SHEETS.ACTIVITIES);
  if (!sh) throw new Error("logActivity_: ACTIVITIES sheet not found");

  const actor = Session.getActiveUser().getEmail() || "unknown";
  const now = nowIso_();

  const rowObj = {
    ACTIVITY_ID: makeId_("ACT"),
    TS: now,
    ACTOR: actor,
    CLIENT_ID: p.clientId || "",
    MATTER_ID: p.matterId || "",
    ACTION: p.action || "",
    MESSAGE: p.message || "",
    META_JSON: json_(p.meta || {}),
  };

  appendRowByHeaders_(sh, c.HEADERS.ACTIVITIES, rowObj);
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
      if (idxUpdated !== -1) sh.getRange(rowNumber, idxUpdated + 1).setValue(nowIso);
      return;
    }
  }
}