/** crmClientCards.gs */

function crm_getClientCard(clientId, opt) {
  opt = opt || {};

  const limitTasks = Number(opt.limitTasks || 20);
  const limitActivities = Number(opt.limitActivities || 20);
  const limitDocuments = Number(opt.limitDocuments || 20);
  const tasksStatus = opt.tasksStatus === undefined ? "OPEN" : opt.tasksStatus;

  if (!clientId) throw new Error("crm_getClientCard: missing clientId");

  const client = crm_findClientById(clientId);
  if (!client) {
    return { ok: false, error: "CLIENT_NOT_FOUND", clientId };
  }

  const matters = crm_findMattersByClientId(clientId) || [];

  const tasksOpen = crm_listTasks({
    clientId: clientId,
    status: "OPEN",
    limit: limitTasks,
  });

  const tasksAll = crm_listTasks({
    clientId: clientId,
    status: "",
    limit: 500,
  });

  const tasksOther = tasksAll.filter(function(t) {
    return String(t.STATUS || "").toUpperCase() !== "OPEN";
  });

  const activities = crm_listActivities({
    clientId: clientId,
    limit: limitActivities,
  });

  const documents = crm_listDocumentsByClientId(clientId, {
    limit: limitDocuments,
  });

  return {
    ok: true,
    client,
    matters,
    tasks: tasksOpen,
    tasksOpen,
    tasksOther,
    activities,
    documents,
    meta: {
      mattersCount: matters.length,
      tasksCount: tasksOpen.length,
      activitiesCount: activities.length,
      documentsCount: documents.length,
    },
  };
}