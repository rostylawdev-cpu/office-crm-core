/** crmCards.gs */

function crm_getMatterCard(matterId, opt) {
  opt = opt || {};

  const limitTasks = opt.limitTasks ?? 20;
  const limitActivities = opt.limitActivities ?? 20;
  const limitDocuments = opt.limitDocuments ?? 20;
  const tasksStatus = opt.tasksStatus ?? "OPEN"; // null = все

  if (!matterId) throw new Error("crm_getMatterCard: missing matterId");

  const matter = crm_getMatterById(matterId);
  if (!matter) {
    return { ok: false, error: "MATTER_NOT_FOUND", matterId };
  }

  const client = matter.CLIENT_ID ? crm_findClientById(matter.CLIENT_ID) : null;

  const tasksOpen = crm_listTasks({
    matterId: matterId,
    status: "OPEN",
    limit: limitTasks,
  });

  const tasksAll = crm_listTasks({
    matterId: matterId,
    status: "",
    limit: 500,
  });

  const tasksOther = tasksAll.filter(function(t) {
    return String(t.STATUS || "").toUpperCase() !== "OPEN";
  });

  const activities = crm_listActivities({
    matterId: matterId,
    limit: limitActivities,
  });

  const documents = crm_listDocumentsByMatterId(matterId, {
    limit: limitDocuments,
  });

  return {
    ok: true,
    client,
    matter,
    tasks: tasksOpen,
    tasksOpen,
    tasksOther,
    activities,
    documents,
  };
}