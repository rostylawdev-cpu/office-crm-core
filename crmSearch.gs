/** crmSearch.gs */

function crm_searchAll(query, opt) {
  opt = opt || {};

  const q = String(query || "").trim();
  const qNorm = q.toLowerCase();
  const limit = Number(opt.limit || 10);

  if (!q) {
    return {
      ok: true,
      query: q,
      leads: [],
      clients: [],
      matters: [],
      meta: {
        leadsCount: 0,
        clientsCount: 0,
        mattersCount: 0,
      },
    };
  }

  const leads = crm_searchLeads_(q, qNorm, limit);
  const clients = crm_searchClients_(q, qNorm, limit);
  const matters = crm_searchMatters_(q, qNorm, limit);

  return {
    ok: true,
    query: q,
    leads,
    clients,
    matters,
    meta: {
      leadsCount: leads.length,
      clientsCount: clients.length,
      mattersCount: matters.length,
    },
  };
}

function crm_searchLeads_(q, qNorm, limit) {
  const rows = crm_listLeads({ limit: 500 }) || [];
  const qPhone = normPhone_(q);

  const out = rows.filter(function (r) {
    return (
      matchField_(r.LEAD_ID, q, qNorm) ||
      matchField_(r.FULL_NAME, q, qNorm) ||
      matchField_(r.FULL_NAME_RU, q, qNorm) ||
      matchField_(r.FULL_NAME_HE, q, qNorm) ||
      matchField_(r.EMAIL, q, qNorm) ||
      matchField_(r.PHONE, q, qNorm, qPhone) ||
      matchField_(r.CASE_TYPE, q, qNorm) ||
      matchField_(r.SOURCE, q, qNorm)
    );
  });

  return out.slice(0, limit);
}

function crm_searchClients_(q, qNorm, limit) {
  const rows = crm_getAllRowsFromSheet_(cfg_().SHEETS.CLIENTS, cfg_().HEADERS.CLIENTS) || [];
  const qPhone = normPhone_(q);

  const out = rows.filter(function (r) {
    return (
      matchField_(r.CLIENT_ID, q, qNorm) ||
      matchField_(r.FULL_NAME, q, qNorm) ||
      matchField_(r.EMAIL, q, qNorm) ||
      matchField_(r.PHONE, q, qNorm, qPhone) ||
      matchField_(r.ID_NUMBER, q, qNorm) ||
      matchField_(r.SOURCE, q, qNorm)
    );
  });

  return out.slice(0, limit);
}

function crm_searchMatters_(q, qNorm, limit) {
  const rows = crm_getAllRowsFromSheet_(cfg_().SHEETS.MATTERS, cfg_().HEADERS.MATTERS) || [];

  const out = rows.filter(function (r) {
    return (
      matchField_(r.MATTER_ID, q, qNorm) ||
      matchField_(r.TITLE, q, qNorm) ||
      matchField_(r.CATEGORY, q, qNorm) ||
      matchField_(r.STAGE, q, qNorm) ||
      matchField_(r.AUTHORITY, q, qNorm) ||
      matchField_(r.CLIENT_ID, q, qNorm)
    );
  });

  return out.slice(0, limit);
}

function matchField_(value, rawQuery, lowerQuery, normalizedPhoneQuery) {
  const v = String(value || "").trim();
  if (!v) return false;

  const vLower = v.toLowerCase();

  if (vLower.indexOf(lowerQuery) !== -1) return true;
  if (v === rawQuery) return true;

  if (normalizedPhoneQuery) {
    const vp = normPhone_(v);
    if (vp && vp === normalizedPhoneQuery) return true;
  }

  return false;
}