/** crmCards.gs */

function crm_getMatterCard(matterId, opt) {
  opt = opt || {};

  const limitTasks = opt.limitTasks ?? 20;
  const limitActivities = opt.limitActivities ?? 20;
  const tasksStatus = opt.tasksStatus ?? "OPEN"; // null = все

  if (!matterId) throw new Error("crm_getMatterCard: missing matterId");

  // Matter
  const matter = crm_getMatterById(matterId);
  if (!matter) {
    return { ok: false, error: "MATTER_NOT_FOUND", matterId };
  }

  // Client
  const client = matter.CLIENT_ID ? crm_findClientById(matter.CLIENT_ID) : null;

  // Tasks
  const tasks = crm_listTasks({
    matterId: matterId,
    status: tasksStatus,
    limit: limitTasks,
  });

  // Activities
  const activities = crm_listActivities({
    matterId: matterId,
    limit: limitActivities,
  });

  return {
    ok: true,
    client,
    matter,
    tasks,
    activities,
  };
}