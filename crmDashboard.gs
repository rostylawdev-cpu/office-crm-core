function crm_getClientDashboard(clientId) {

  if (!clientId) {
    throw new Error("crm_getClientDashboard: missing clientId");
  }

  const client = crm_findClientById(clientId);

  const matters = crm_findMattersByClientId(clientId);

  const tasks = crm_listTasks({
    clientId: clientId,
    status: "OPEN",
    limit: 50
  });

  const activities = crm_listActivities({
    clientId: clientId,
    limit: 20
  });

  return {
    client,
    matters,
    openTasks: tasks,
    recentActivities: activities
  };
}