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

  // Normalize EVENT_DATE: GAS reads date cells as Date objects.
  // Emit YYYY-MM-DD string (for <input type="date">) and DD.MM.YYYY (for display).
  var eventDateIso = "";
  var eventDateDisplay = "";
  var rawEventDate = matter.EVENT_DATE;
  if (rawEventDate instanceof Date && !isNaN(rawEventDate.getTime())) {
    var y = rawEventDate.getFullYear();
    var mo = String(rawEventDate.getMonth() + 1).padStart(2, "0");
    var d = String(rawEventDate.getDate()).padStart(2, "0");
    eventDateIso = y + "-" + mo + "-" + d;
    eventDateDisplay = d + "." + mo + "." + y;
  } else if (rawEventDate) {
    var s = String(rawEventDate).trim();
    // Already a YYYY-MM-DD string
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
      eventDateIso = s;
      var parts = s.split("-");
      eventDateDisplay = parts[2] + "." + parts[1] + "." + parts[0];
    } else if (/^\d{2}\.\d{2}\.\d{4}$/.test(s)) {
      // Already DD.MM.YYYY
      eventDateDisplay = s;
      var dp = s.split(".");
      eventDateIso = dp[2] + "-" + dp[1] + "-" + dp[0];
    } else {
      eventDateDisplay = s;
      eventDateIso = s;
    }
  }
  matter.EVENT_DATE = eventDateIso;           // YYYY-MM-DD for <input type="date">
  matter.EVENT_DATE_DISPLAY = eventDateDisplay; // DD.MM.YYYY for visible label

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