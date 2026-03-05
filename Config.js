/** Config.js */

const CFG = {
  // Script Property key name (в Script Properties будет лежать сам Spreadsheet ID)
  CRM_SPREADSHEET_ID_PROP: "CRM_SPREADSHEET_ID",

  SHEETS: {
    CLIENTS: "Clients",
    MATTERS: "Matters",
    DOCUMENTS: "Documents",
    ACTIVITIES: "Activities",
    TASKS: "Tasks",
  },

  HEADERS: {
    CLIENTS: [
      "CLIENT_ID", "FULL_NAME", "PHONE", "EMAIL", "ID_NUMBER", "LOCALE", "STATUS", "OWNER", "SOURCE",
      "FOLDER_URL", "CREATED_AT", "UPDATED_AT", "LAST_ACTIVITY_AT",
      "PAYMENT_STATUS", "PAYMENT_DUE", "PAYMENT_PAID", "PAYMENT_NOTE",
    ],
    MATTERS: [
      "MATTER_ID", "CLIENT_ID", "CATEGORY", "TITLE", "STAGE", "OWNER", "AUTHORITY",
      "FOLDER_URL", "OPENED_AT", "CLOSED_AT", "UPDATED_AT", "LAST_ACTIVITY_AT", "SUMMARY_SHORT",
    ],
    DOCUMENTS: [
      "DOC_ID", "CLIENT_ID", "MATTER_ID", "TYPE", "STATUS", "TITLE", "DOC_URL", "PDF_URL", "FILE_ID",
      "CREATED_AT", "CREATED_BY", "NOTES",
    ],
    ACTIVITIES: [
      "ACTIVITY_ID", "TS", "ACTOR", "CLIENT_ID", "MATTER_ID", "ACTION", "MESSAGE", "META_JSON",
    ],
    TASKS: [
      "TASK_ID", "CLIENT_ID", "MATTER_ID", "TYPE", "TITLE", "DUE_DATE", "STATUS", "PRIORITY",
      "GENERATED_BY", "ASSIGNEE", "CREATED_AT", "DONE_AT", "NOTES",
    ],
  },
};

function cfg_() { return CFG; }