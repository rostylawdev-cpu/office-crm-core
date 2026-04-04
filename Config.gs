/** Config.gs */

const CFG = {
  // Script Property key name (в Script Properties хранится Spreadsheet ID)
  CRM_SPREADSHEET_ID_PROP: "CRM_SPREADSHEET_ID",

  SHEETS: {
    LEADS: "Leads",
    CLIENTS: "Clients",
    MATTERS: "Matters",
    DOCUMENTS: "Documents",
    ACTIVITIES: "Activities",
    TASKS: "Tasks",
  },

    DRIVE: {
    // папка хранения CRM-документов
    // если пусто — файлы будут падать в корень Drive владельца скрипта
    ROOT_FOLDER_ID: "1lscoZRCKKaTRgivFIPduZHWLcYXYvyYZ"
  },

  HEADERS: {
    LEADS: [
      "LEAD_ID",
      "CREATED_AT",
      "UPDATED_AT",
      "SOURCE",
      "CAMPAIGN",
      "FULL_NAME",
      "PHONE",
      "EMAIL",
      "CASE_TYPE",
      "DESCRIPTION",
      "STATUS",
      "ASSIGNED_TO",
      "CLIENT_ID",
      "NOTES",
      "ID_TYPE",
      "ID_NUMBER",
      "ADDRESS",
    ],

    CLIENTS: [
      "CLIENT_ID",
      "FULL_NAME",
      "PHONE",
      "EMAIL",
      "ID_NUMBER",
      "LOCALE",
      "STATUS",
      "OWNER",
      "SOURCE",
      "FOLDER_URL",
      "CREATED_AT",
      "UPDATED_AT",
      "LAST_ACTIVITY_AT",
      "PAYMENT_STATUS",
      "PAYMENT_DUE",
      "PAYMENT_PAID",
      "PAYMENT_NOTE",
      "ID_TYPE",
      "ADDRESS",
    ],

    MATTERS: [
      "MATTER_ID",
      "CLIENT_ID",
      "CATEGORY",
      "TITLE",
      "STAGE",
      "OWNER",
      "AUTHORITY",
      "FOLDER_URL",
      "OPENED_AT",
      "CLOSED_AT",
      "UPDATED_AT",
      "LAST_ACTIVITY_AT",
      "SUMMARY_SHORT",
    ],

    DOCUMENTS: [
      "DOC_ID",
      "CLIENT_ID",
      "MATTER_ID",
      "TYPE",
      "STATUS",
      "TITLE",
      "DOC_URL",
      "PDF_URL",
      "FILE_ID",
      "CREATED_AT",
      "CREATED_BY",
      "NOTES",
    ],

    ACTIVITIES: [
      "ACTIVITY_ID",
      "TS",
      "ACTOR",
      "CLIENT_ID",
      "MATTER_ID",
      "ACTION",
      "MESSAGE",
      "META_JSON",
    ],

    TASKS: [
      "TASK_ID",
      "CLIENT_ID",
      "MATTER_ID",
      "TYPE",
      "TITLE",
      "DUE_DATE",
      "STATUS",
      "PRIORITY",
      "GENERATED_BY",
      "ASSIGNEE",
      "CREATED_AT",
      "UPDATED_AT",
      "DONE_AT",
      "NOTES",
    ],
  },

  TASK_TEMPLATES: {
    WORK_ACCIDENT: [
      { type: "DOC_REQUEST", title: "Запросить טופס 250 + תלושים + דו\"חות נוכחות", days: 0, priority: "HIGH" },
      { type: "DOC_REQUEST", title: "Запросить медицинские документы + סיכום ביקור", days: 0, priority: "HIGH" },
      { type: "FORM_PREP", title: "Подготовить/проверить טופס בל/211", days: 1, priority: "HIGH" },
      { type: "FOLLOW_UP", title: "Созвон с клиентом: уточнить обстоятельства התאונה", days: 1, priority: "MEDIUM" },
      { type: "SUBMIT", title: "Подача в БЛ + контроль подтверждения", days: 2, priority: "HIGH" },
    ],

    LABOR_DISPUTE: [
      { type: "DOC_REQUEST", title: "Запросить תלושים/דוחות נוכחות/חוזה עבודה", days: 0, priority: "HIGH" },
      { type: "CALC", title: "Сделать первичный расчет (שכר/פנסיה/פיצויים)", days: 1, priority: "HIGH" },
      { type: "LETTER", title: "Подготовить מכתב התראה", days: 2, priority: "MEDIUM" },
    ],
  },
};

function cfg_() {
  return CFG;
}