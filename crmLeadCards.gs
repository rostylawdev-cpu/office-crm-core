/** crmLeadCards.gs */

function crm_getLeadCard(leadId, opt) {
  opt = opt || {};

  const limitActivities = Number(opt.limitActivities || 20);

  if (!leadId) throw new Error("crm_getLeadCard: missing leadId");

  const lead = crm_getLeadById(leadId);
  if (!lead) {
    return { ok: false, error: "LEAD_NOT_FOUND", leadId };
  }

  const client = lead.CLIENT_ID ? crm_findClientById(lead.CLIENT_ID) : null;

  const matters = lead.CLIENT_ID
    ? crm_findMattersByClientId(lead.CLIENT_ID)
    : [];

  const activities = crm_listLeadActivities_(lead, limitActivities);

  return {
    ok: true,
    lead,
    client,
    matters,
    activities,
    meta: {
      mattersCount: matters.length,
      activitiesCount: activities.length,
      linkedClient: !!client,
    },
  };
}

/**
 * Собираем активности для карточки лида.
 * В текущей модели ACTIVITIES не имеют LEAD_ID,
 * поэтому показываем:
 * 1) активности клиента, если лид уже конвертирован
 * 2) иначе — пусто
 */
function crm_listLeadActivities_(lead, limit) {
  if (!lead || !lead.CLIENT_ID) return [];

  return crm_listActivities({
    clientId: lead.CLIENT_ID,
    limit: limit || 20,
  });
}