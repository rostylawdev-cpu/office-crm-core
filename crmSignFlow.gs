/** crmSignFlow.gs */

const CRM_SIGN_TOKEN_PREFIX = "CRM_SIGN_TOKEN_";
const CRM_SIGN_TTL_HOURS = 72;
const CRM_SIGN_MARKER = "[[SIGN_HERE]]";

const CRM_SIGN_TOKEN_STATUS_ACTIVE = "ACTIVE";
const CRM_SIGN_TOKEN_STATUS_USED = "USED";
const CRM_SIGN_TOKEN_STATUS_EXPIRED = "EXPIRED";

function crm_getSignTokenKey(token) {
  return CRM_SIGN_TOKEN_PREFIX + token;
}

function crm_getSignTokenEntry(token) {
  if (!token) return null;
  const raw = PropertiesService.getScriptProperties().getProperty(crm_getSignTokenKey(token));
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function crm_setSignTokenEntry(token, entry) {
  if (!token || !entry) throw new Error("crm_setSignTokenEntry: token and entry are required");
  PropertiesService.getScriptProperties().setProperty(crm_getSignTokenKey(token), JSON.stringify(entry));
}

function crm_expireSignToken(token) {
  const entry = crm_getSignTokenEntry(token);
  if (!entry) return null;
  entry.status = CRM_SIGN_TOKEN_STATUS_EXPIRED;
  entry.expiresAt = new Date().toISOString();
  crm_setSignTokenEntry(token, entry);
  return entry;
}

function crm_markSignTokenUsed(token) {
  const entry = crm_getSignTokenEntry(token);
  if (!entry) throw new Error("crm_markSignTokenUsed: token not found");
  entry.status = CRM_SIGN_TOKEN_STATUS_USED;
  entry.usedAt = new Date().toISOString();
  crm_setSignTokenEntry(token, entry);
  return entry;
}

function crm_isTokenExpired(entry) {
  if (!entry || !entry.expiresAt) return true;
  const expires = new Date(entry.expiresAt).getTime();
  return Number.isNaN(expires) || expires < Date.now();
}

function crm_findActiveSignTokenForDoc(docId) {
  if (!docId) return null;
  const all = PropertiesService.getScriptProperties().getProperties();
  const prefix = CRM_SIGN_TOKEN_PREFIX;
  for (const key in all) {
    if (!key.startsWith(prefix)) continue;
    if (key.endsWith("_expires")) continue;

    const raw = all[key];
    if (!raw) continue;
    let entry;
    try {
      entry = JSON.parse(raw);
    } catch (e) {
      continue;
    }
    if (!entry || entry.docId !== docId) continue;
    if (entry.kind === "ONBOARDING_PACKAGE") continue; // package tokens are not per-doc tokens

    if (entry.status !== CRM_SIGN_TOKEN_STATUS_ACTIVE) continue;

    if (crm_isTokenExpired(entry)) {
      entry.status = CRM_SIGN_TOKEN_STATUS_EXPIRED;
      crm_setSignTokenEntry(entry.token, entry);
      continue;
    }

    return entry;
  }
  return null;
}

function crm_getLeadIdByClientId(clientId) {
  if (!clientId) return null;
  const leads = crm_listLeads({ limit: 500 });
  return (leads || []).find((l) => String(l.CLIENT_ID || "").trim() === String(clientId).trim())?.LEAD_ID || null;
}

function crm_safelyConvertLeadForClient(clientId) {
  if (!clientId) return null;
  const leadId = crm_getLeadIdByClientId(clientId);
  if (!leadId) return null;

  const lead = crm_getLeadById(leadId);
  if (!lead) return null;
  if (String(lead.STATUS || "").toUpperCase() === "CONVERTED") return { ok: true, leadId, converted: false, alreadyConverted: true };

  // crm_getLeadIdByClientId only finds leads where CLIENT_ID is already set,
  // so crm_convertLeadToClient would short-circuit with alreadyLinked:true
  // without ever writing STATUS. Finalize conversion by updating STATUS directly.
  crm_updateLeadStatus(leadId, "CONVERTED");
  return { ok: true, leadId, clientId: lead.CLIENT_ID, converted: true };
}

const CRM_TEMPLATE_AGREEMENT = `הסכם ייצוג ושכ"ט / Соглашение о представительстве и оплате труда адвоката

ДАННЫЕ СТОРОН
פרטי הצדדים

Адвокатский офис:
{{OFFICE_NAME_RU}}
Адрес: {{OFFICE_ADDRESS_RU}}
Тел.: {{OFFICE_PHONE}}
E-mail: {{OFFICE_EMAIL}}

עוה"ד:
{{OFFICE_NAME_HE}}
כתובת: {{OFFICE_ADDRESS_HE}}
טל’: {{OFFICE_PHONE}}
דוא"ל: {{OFFICE_EMAIL}}

Клиент: {{CLIENT_FULL_NAME_RU}}
Документ: {{ID_DOC_RU}}
Адрес: {{ADDRESS_RU}}
Тел.: {{PHONE}}

הלקוח: {{CLIENT_FULL_NAME_HE}}
מסמך מזהה: {{ID_DOC_HE}}
כתובת: {{ADDRESS_HE}}
טל’: {{PHONE}}

Ввиду того, что Клиент обратился в Адвокатское бюро с просьбой заняться его делом, Адвокатское бюро соглашается взять на себя ведение дела Клиента при соблюдении условий настоящего соглашения.
הואיל והלקוח פנה אל משרד עוה"ד בבקשה לטפל בעניינו, ומשרד עוה"ד הסכים ליטול על עצמו את הטיפול והייצוג בכפוף לתנאי הסכם זה.

1. Суть вопроса:
{{SUBJECT_RU}}
1. מהות העניין:
{{SUBJECT_HE}}

2. Общие условия:
2.1. Адвокат обязуется добросовестно и профессионально выполнять обязательства.
2.2. Деятельность Адвоката не гарантирует конкретного результата.
2.3. Клиент обязуется предоставлять полную и достоверную информацию.
2.4. В случае введения в заблуждение Адвокат вправе прекратить представительство.
2.5. Клиент не вправе вести переговоры без участия Адвоката.
2.6. Существенное нарушение условий является основанием для расторжения.
2. תנאים כלליים:
2.1. עוה"ד יפעלו בנאמנות ובמקצועיות.
2.2. אין התחייבות לתוצאה מסוימת.
2.3. הלקוח ימסור מידע מלא ונכון.
2.4. הטעיה תאפשר הפסקת ייצוג.
2.5. הלקוח לא ינהל מו"מ ללא עוה"ד.
2.6. הפרה יסודית תאפשר ביטול ההסכם.

3. Гонорар:
{{PAYMENT_RU}}
3. שכר טרחה:
{{PAYMENT_HE}}

4. Расходы:
Госпошлины и иные расходы оплачиваются Клиентом отдельно. Апелляции не включены.
4. הוצאות:
אגרות והוצאות ישולמו בנפרד. ערעורים אינם כלולים.

Подпись клиента / חתימה:

[[SIGN_HERE]]

Подпись адвоката / חתימה עו"ד:


Дата: {{DATE_SIGN}}`;

const CRM_TEMPLATE_POA = `ייפוי כח / Доверенность

{{POA_MANDATE_RU}}

{{POA_MANDATE_HE}}

1. Подписывать и подавать любые иски или встречные иски и/или любые ходатайства, возражения, заявления о защите, возражения, апелляции, уведомления, или любые иные процессуальные действия.
2. Запрашивать и получать медицинскую документацию у любого врача и/или медицинского учреждения.
3. Запрашивать и получать медицинские и/или профессиональные заключения у любого врача и/или медицинского учреждения.
4. Выступать в связи с любым из вышеуказанных действий перед всеми судами, судебными инстанциями.
5. Передавать любой вопрос, относящийся к вышеуказанным делам, на рассмотрение арбитража и подписывать арбитражное соглашение.
6. Заключать мировые соглашения по любым вопросам, относящимся или вытекающим из вышеуказанных дел.
7. Взыскивать сумму иска или любую иную сумму по любому из вышеуказанных дел, включая судебные расходы и вознаграждение адвоката, получать от моего имени любые документы и имущество.
8. Осуществлять любые действия и подписывать любые документы или письменные акты без исключения.
9. Подписывать от моего имени и вместо меня любые ходатайства, заявления и иные документы, давать заявления, расписки и подтверждения, а также получать любые документы.
10. Передавать настоящую доверенность со всеми содержащимися в ней полномочиями либо с их частью другому адвокату, с правом дальнейшей передачи полномочий третьим лицам, увольнять таких лиц и назначать других вместо них, вести вышеуказанные дела по своему усмотрению и, в целом, совершать все действия, которые он сочтёт необходимыми и полезными в связи с судебным процессом или вышеуказанными делами; я заранее подтверждаю законную силу всех действий, совершённых моим представителем или его заместителями на основании настоящей доверенности.

1. לחתום על ולהגיש כל תביעה או תביעה שכנגד, ו/או כל בקשה, הגנה, התנגדות, בקשה למתן רשות לערער, ערעור, הודעה, טענה, תובענה או כל הליך אחר הנוגע או הנובע מההליך הנ"ל ללא יוצא מן הכלל ומבלי לפגוע באמור גם להודות ו/או לכפור בשמי במשפטים פליליים.
2. לבקש ולקבל רשומה רפואית מכל רופא ו/או מוסד רפואי שבדק אותי ו/או טיפל בי והכל בהתאם לכתב ויתור על סודיות רפואית.
3. לבקש ולקבל חוות דעת רפואית ו/או מקצועית מכל רופא ו/או מוסד רפואי ו/או כל בעל מקצוע אחר.
4. להופיע בקשר לכל אחת מהפעולות הנ"ל בפני כל בתי המשפט, בתי דין למיניהם או מוסדות אחרים הן ממשלתיים והן אחרים עד לדרגה האחרונה.
5. למסור כל עניין הנוגע מהעניין האמור לעיל לבוררות ולחתום של שטר בוררין כפי שבא כחי ימצא לנכון ולמועיל.
6. להתפשר בכל עניין הנוגע או הנובע מהעניינים האמורים לעיל לפי שקול דעתו של בא כוחי ולחתום על פשרה כזו בביהמ"ש או מחוצה לו.
7. לגבות את סכום התביעה או כל סכום אחר בכל עניין מהעניינים הנ"ל לרבות הוצאות בימ"ש ושכר טרחת עו"ד, לקבל בשמי כל מסמך וחפץ ולתת קבלות ושחרורים כפי שבא כחי ימצא לנכון ולמתאים.
8. לנקוט בכל הפעולות ולחתום על כל מסמך או כתב בלי יוצא מן הכלל אשר בא כוחי ימצא לנכון בכל עניין הנובע מהעניין הנ"ל.
9. לחתום בשמי במקומי על כל בקשה, הצהרה ומסמכים אחרים למיניהם וליתן הצהרות, קבלות ואישורים ולקבל כל מסמך שאני רשאי לקבלו עפ"י דין.
10. להעביר יפוי כח זה על כל הסמכויות שבו או חלק מהן לעו"ד אחר עם זכות העברה לאחרים, לפטרם ולמנות אחרים במקומם ולנהל את עניני הנ"ל לפי ראות עיניו ובכלל לעשות את כל הצעדים שימצא לנכון ומועיל בקשר עם המשפט או עם עניני הנ"ל מאשר את מעשיו או מעשי ממלאי המקום בתוקף יפוי כח זה מראש.

Подпись клиента / חתימה:

[[SIGN_HERE]]

Подпись адвоката / חתימה עו"ד:________________________________________
{{LAWYER_CONFIRM_RU}} / {{LAWYER_CONFIRM_HE}}`;

// REQUIRED Script Properties for legal document generation.
// Set via Apps Script editor → Project Settings → Script Properties.
// Without these the POA mandate sentence will have placeholders/blanks:
//   OFFICE_NAME_RU        — office name in Russian  (e.g. "Адвокатское бюро Иванова")
//   OFFICE_NAME_HE        — office name in Hebrew   (e.g. "משרד עו\"ד כהן")
//   OFFICE_ADDRESS        — office address in Russian (also read as OFFICE_ADDRESS_RU)
//   OFFICE_ADDRESS_HE     — office address in Hebrew
//   OFFICE_PHONE          — phone number
//   OFFICE_EMAIL          — email address
//   LAWYERS_RU            — lawyer names in Russian (e.g. "А. Иванов, М. Петрова") ← CRITICAL for POA
//   LAWYERS_HE            — lawyer names in Hebrew  (e.g. "א. כהן, מ. לוי")           ← CRITICAL for POA
//   LAW_OFFICE_RU         — firm name in Russian (defaults to OFFICE_NAME_RU if not set)
//   LAW_OFFICE_HE         — firm name in Hebrew  (defaults to OFFICE_NAME_HE if not set)
//   LAWYER_CONFIRM_RU     — optional: confirmation line (has safe default)
//   LAWYER_CONFIRM_HE     — optional: confirmation line (has safe default)
function crm_getOfficeProfile() {
  const props = PropertiesService.getScriptProperties();
  const nameRu = String(props.getProperty("OFFICE_NAME_RU") || props.getProperty("OFFICE_NAME") || "").trim();
  return {
    OFFICE_NAME: nameRu,
    OFFICE_ADDRESS: String(props.getProperty("OFFICE_ADDRESS") || props.getProperty("OFFICE_ADDRESS_RU") || "").trim(),
    OFFICE_PHONE: String(props.getProperty("OFFICE_PHONE") || "").trim(),
    OFFICE_EMAIL: String(props.getProperty("OFFICE_EMAIL") || "").trim(),
    OFFICE_NAME_HE: String(props.getProperty("OFFICE_NAME_HE") || "").trim(),
    OFFICE_ADDRESS_HE: String(props.getProperty("OFFICE_ADDRESS_HE") || "").trim(),
    LAWYERS_RU: String(props.getProperty("LAWYERS_RU") || "").trim(),
    LAWYERS_HE: String(props.getProperty("LAWYERS_HE") || "").trim(),
    LAW_OFFICE_RU: String(props.getProperty("LAW_OFFICE_RU") || nameRu).trim(),
    LAW_OFFICE_HE: String(props.getProperty("LAW_OFFICE_HE") || props.getProperty("OFFICE_NAME_HE") || "").trim(),
    LAWYER_CONFIRM_RU: String(props.getProperty("LAWYER_CONFIRM_RU") || "Подпись подтверждает согласие с условиями договора.").trim(),
    LAWYER_CONFIRM_HE: String(props.getProperty("LAWYER_CONFIRM_HE") || "החתימה מאשרת הסכמה לתנאי ההסכם.").trim(),
  };
}

function crm_buildTemplateData(client, matter) {
  const office = crm_getOfficeProfile();

  const safeClientIdNumber = String(client.ID_NUMBER || "").trim();
  const clientIdNumberFinal = safeClientIdNumber && !safeClientIdNumber.startsWith("LEAD_") ? safeClientIdNumber : "";

  const isWorkAccident = String(matter.CATEGORY || "").toUpperCase() === "WORK_ACCIDENT";
  const authority = String(matter.AUTHORITY || "").trim();

  // Map matter.AUTHORITY to language-specific display strings.
  // Handles enum values from the form (LABOR_COURT, LABOR_INSURANCE, …),
  // the "Bituach Leumi" string written by onboarding init, and free-text values.
  // If authority is empty, work-accident default fills each language separately.
  var authorityNorm = authority.toUpperCase().replace(/[\s-]+/g, "_");
  var authorityHe, authorityRu;
  if (authorityNorm === "LABOR_COURT") {
    authorityHe = "בית הדין לעבודה";
    authorityRu = "суд по трудовым спорам";
  } else if (authorityNorm === "NATIONAL_LABOR_COURT") {
    authorityHe = "בית הדין הארצי לעבודה";
    authorityRu = "Национальный суд по трудовым спорам";
  } else if (authorityNorm === "LABOR_INSURANCE" || authorityNorm === "BITUACH_LEUMI" || authorityNorm === "BITUACH LEUMI") {
    authorityHe = "המוסד לביטוח לאומי";
    authorityRu = "Битуах Леуми";
  } else if (authorityNorm === "EMPLOYER") {
    authorityHe = "המעסיק";
    authorityRu = "работодатель";
  } else if (authority === "") {
    // Empty: apply work-accident language-specific defaults; otherwise leave blank
    authorityHe = isWorkAccident ? "המוסד לביטוח לאומי" : "";
    authorityRu = isWorkAccident ? "Битуах Леуми" : "";
  } else {
    // Free-text entered by staff — pass through as-is to both (staff knows the language)
    authorityHe = authority;
    authorityRu = authority;
  }

  // Map raw client.ID_TYPE to language-appropriate display values.
  var rawIdType = String(client.ID_TYPE || "").trim();
  var idTypeNorm = rawIdType.toUpperCase().replace(/[\s_-]+/g, "");
  var idTypeRu, idTypeHe;
  if (idTypeNorm === "TZ" || idTypeNorm === "TEUDATZEHUT") {
    idTypeRu = "теудат зеут";
    idTypeHe = "תעודת זהות";
  } else if (idTypeNorm === "PASSPORT" || idTypeNorm === "DARKON") {
    idTypeRu = "паспорт";
    idTypeHe = "דרכון";
  } else if (idTypeNorm === "OTHER") {
    idTypeRu = "документ";
    idTypeHe = "מסמך מזהה";
  } else if (rawIdType === "") {
    idTypeRu = "";
    idTypeHe = "";
  } else {
    // Unknown type — pass through raw value rather than produce wrong-language text
    idTypeRu = rawIdType;
    idTypeHe = rawIdType;
  }

  // Build BL_CONTEXT_HE from matter category + stored event date.
  // The event date is never derived from lead creation date; it comes from matter.EVENT_DATE.
  var blContextHe = "";
  if (isWorkAccident) {
    var rawEventDate = String(matter.EVENT_DATE || "").trim();
    var formattedEventDate = "";
    if (rawEventDate) {
      try {
        var evtD = new Date(rawEventDate);
        if (!isNaN(evtD.getTime())) {
          formattedEventDate = ("0" + evtD.getDate()).slice(-2) + "." +
                               ("0" + (evtD.getMonth() + 1)).slice(-2) + "." +
                               evtD.getFullYear();
        }
      } catch (e) {}
      if (!formattedEventDate) formattedEventDate = rawEventDate; // raw fallback if Date() fails
    }
    blContextHe = formattedEventDate
      ? "(בקשר לתאונת עבודה מיום " + formattedEventDate + ", לרבות זכויות מול המוסד לביטוח לאומי)"
      : "(בקשר לתאונת עבודה, לרבות זכויות מול המוסד לביטוח לאומי)";
  }

  // ── Safe composites ────────────────────────────────────────────────────────
  // These collapse empty ID_TYPE / lawyer / authority values so the output does
  // not contain dangling commas or fragments like "документ: ," or "в .".

  // ID_DOC_RU / ID_DOC_HE: joins ID_TYPE label + ID_NUMBER, skipping empties.
  var idDocRu = [idTypeRu, clientIdNumberFinal].filter(Boolean).join(" ");
  var idDocHe = [idTypeHe, clientIdNumberFinal].filter(Boolean).join(" ");

  // POA_MANDATE_RU: full Russian mandate sentence, safe when office props missing.
  var poaMandateRu = (function () {
    var nameP = String(client.FULL_NAME || "").trim();
    var idP   = idDocRu ? ", документ: " + idDocRu : "";
    var lawyersOfficeP = "";
    if (office.LAWYERS_RU && office.LAW_OFFICE_RU) {
      lawyersOfficeP = "адвокатов: " + office.LAWYERS_RU +
        ", действующих от имени адвокатского офиса " + office.LAW_OFFICE_RU;
    } else if (office.LAWYERS_RU) {
      lawyersOfficeP = "адвокатов: " + office.LAWYERS_RU;
    } else if (office.LAW_OFFICE_RU) {
      lawyersOfficeP = "адвокатского офиса " + office.LAW_OFFICE_RU;
    }
    var mandateP = lawyersOfficeP ? "настоящим уполномочиваю " + lawyersOfficeP : "настоящим уполномочиваю";
    var authorityP = authorityRu ? ", представлять мои интересы в " + authorityRu : "";
    return "Я, нижеподписавшийся " + nameP + idP + ", " + mandateP + authorityP + ".";
  })();

  // POA_MANDATE_HE: full Hebrew mandate sentence, safe when office props missing.
  var poaMandateHe = (function () {
    var nameP  = String(client.FULL_NAME || "").trim();
    var idNumP = clientIdNumberFinal ? ', ת"ז ' + clientIdNumberFinal : "";
    var lawyersOfficeP = "";
    if (office.LAWYERS_HE && office.LAW_OFFICE_HE) {
      lawyersOfficeP = "ממנה בזה את עורכי הדין: " + office.LAWYERS_HE + " ממשרד " + office.LAW_OFFICE_HE;
    } else if (office.LAWYERS_HE) {
      lawyersOfficeP = "ממנה בזה את עורכי הדין: " + office.LAWYERS_HE;
    } else if (office.LAW_OFFICE_HE) {
      lawyersOfficeP = "ממנה בזה ממשרד " + office.LAW_OFFICE_HE;
    } else {
      lawyersOfficeP = "ממנה בזה";
    }
    var authP   = authorityHe ? "ב" + authorityHe : "";
    var authBlP = [authP, blContextHe].filter(Boolean).join(" ");
    var representP = "להיות באי כוחי" + (authBlP ? " " + authBlP : "");
    return 'אני הח"מ ' + nameP + idNumP + ", " + lawyersOfficeP + ", " + representP + ".";
  })();

  return {
    // No fake fallback — empty string is safer than "Client" in a legal document.
    CLIENT_FULL_NAME: String(client.FULL_NAME || "").trim(),
    CLIENT_ID_NUMBER: clientIdNumberFinal,
    CLIENT_ID_TYPE: String(client.ID_TYPE || "").trim(),
    CLIENT_PHONE: String(client.PHONE || "").trim(),
    CLIENT_EMAIL: String(client.EMAIL || "").trim(),

    // Real template placeholders
    ID_NUMBER: clientIdNumberFinal,
    PHONE: String(client.PHONE || "").trim(),

    // No fake fallback — empty title is safer than "Subject matter" in a legal document.
    MATTER_TITLE: String(matter.TITLE || "").trim(),
    MATTER_ID: String(matter.MATTER_ID || "").trim(),
    MATTER_AUTHORITY: authorityHe,
    MATTER_SUBJECT: String(matter.SUMMARY_SHORT || matter.TITLE || "").trim(),

    SUBJECT_RU: String(matter.SUMMARY_SHORT || matter.TITLE || matter.CATEGORY || "").trim(),
    SUBJECT_HE: String(matter.SUMMARY_SHORT_HE || matter.TITLE || matter.CATEGORY || "").trim(),

    AUTHORITY_RU: authorityRu,
    AUTHORITY_HE: authorityHe,

    DATE_SIGN: nowIso_(),
    PAYMENT_TERMS: String(matter.PAYMENT_TERMS || "").trim(),
    PAYMENT_RU: String(matter.PAYMENT_TERMS || matter.PAYMENT_RU || "Согласно отдельной договорённости об оплате.").trim(),
    PAYMENT_HE: String(matter.PAYMENT_HE || "בהתאם להסכם שכר טרחה נפרד.").trim(),

    OFFICE_NAME: office.OFFICE_NAME,
    OFFICE_ADDRESS: office.OFFICE_ADDRESS,
    OFFICE_PHONE: office.OFFICE_PHONE,
    OFFICE_EMAIL: office.OFFICE_EMAIL,

    OFFICE_NAME_RU: office.OFFICE_NAME,
    OFFICE_ADDRESS_RU: office.OFFICE_ADDRESS,
    OFFICE_PHONE_RU: office.OFFICE_PHONE,
    OFFICE_EMAIL_RU: office.OFFICE_EMAIL,
    OFFICE_NAME_HE: office.OFFICE_NAME_HE,
    OFFICE_ADDRESS_HE: office.OFFICE_ADDRESS_HE,

    CLIENT_FULL_NAME_RU: String(client.FULL_NAME || "").trim(),
    CLIENT_FULL_NAME_HE: String(client.FULL_NAME || "").trim(),
    ID_TYPE_RU: idTypeRu,
    ID_TYPE_HE: idTypeHe,
    ADDRESS_RU: String(client.ADDRESS || "").trim(),
    ADDRESS_HE: String(client.ADDRESS || "").trim(),

    LAWYERS_RU: office.LAWYERS_RU,
    LAWYERS_HE: office.LAWYERS_HE,
    LAW_OFFICE_RU: office.LAW_OFFICE_RU,
    LAW_OFFICE_HE: office.LAW_OFFICE_HE,

    LAWYER_CONFIRM_RU: office.LAWYER_CONFIRM_RU,
    LAWYER_CONFIRM_HE: office.LAWYER_CONFIRM_HE,

    BL_CONTEXT_HE: blContextHe,

    // Composites — used by templates to avoid broken fragments when fields are empty
    ID_DOC_RU: idDocRu,
    ID_DOC_HE: idDocHe,
    POA_MANDATE_RU: poaMandateRu,
    POA_MANDATE_HE: poaMandateHe,
  };
}

function crm_templateFill(template, data) {
  if (!template) return "";
  return String(template).replace(/{{\s*([A-Z0-9_]+)\s*}}/g, function (match, key) {
    return key && data && data.hasOwnProperty(key) ? String(data[key]) : "";
  });
}

function crm_generateAgreementAndPoa(matterId) {
  if (!matterId) throw new Error("crm_generateAgreementAndPoa: matterId is required");

  const matter = crm_getMatterById(matterId);
  if (!matter) throw new Error("Matter not found: " + matterId);

  if (!matter.CLIENT_ID) throw new Error("Matter has no CLIENT_ID");

  // Provisional onboarding: CLIENT_ID may be a leadId (LEAD_xxx) before real client is created after signing.
  var client;
  if (String(matter.CLIENT_ID).startsWith("LEAD_")) {
    const lead_ = crm_getLeadById(matter.CLIENT_ID);
    if (!lead_) throw new Error("Lead not found for provisional onboarding matter: " + matter.CLIENT_ID);
    client = {
      CLIENT_ID: matter.CLIENT_ID,
      FULL_NAME: String(lead_.FULL_NAME || "").trim(),
      PHONE: String(lead_.PHONE || "").trim(),
      EMAIL: String(lead_.EMAIL || "").trim(),
      ID_TYPE: String(lead_.ID_TYPE || "").trim(),
      ID_NUMBER: String(lead_.ID_NUMBER || "").trim(),
      ADDRESS: String(lead_.ADDRESS || "").trim(),
      FOLDER_URL: "",
    };
  } else {
    client = crm_findClientById(matter.CLIENT_ID);
    if (!client) throw new Error("Client not found: " + matter.CLIENT_ID);
  }

  // Guard: prevent duplicate AGREEMENT/POA drafts for the same matter
  const existingDocs = crm_listDocumentsByMatterId(matterId, { limit: 100 });
  const hasAgreement = existingDocs.some(function (d) { return (d.TYPE || "").toUpperCase() === "AGREEMENT"; });
  const hasPoa = existingDocs.some(function (d) { return (d.TYPE || "").toUpperCase() === "POA"; });
  if (hasAgreement && hasPoa) {
    crm_logActivity({
      action: "DUPLICATE_DOC_BLOCKED",
      message: "Duplicate Agreement+POA generation blocked for matter " + matterId,
      matterId: matterId,
      clientId: matter.CLIENT_ID || "",
      meta: {},
    });
    return { ok: false, code: "already_generated", message: "Documents already generated for this matter." };
  }

  let matterFolderUrl = matter.FOLDER_URL;
  if (!matterFolderUrl) {
    var clientFolderId = extractFolderIdFromUrl_(client.FOLDER_URL);
    // Provisional onboarding: create a staging folder anchored to the leadId when no client folder exists yet
    if (!clientFolderId && String(matter.CLIENT_ID).startsWith("LEAD_")) {
      const stagingFolderUrl = crm_getOrCreateClientFolder(matter.CLIENT_ID, client.FULL_NAME || matter.CLIENT_ID);
      if (stagingFolderUrl) {
        clientFolderId = extractFolderIdFromUrl_(stagingFolderUrl);
        if (clientFolderId) client = Object.assign({}, client, { FOLDER_URL: stagingFolderUrl });
      }
    }
    if (!clientFolderId) throw new Error("Client folder is not set or invalid");

    const mattersFolder = crm_ensureDriveFolder_(clientFolderId, "02_Matters");
    if (!mattersFolder || !mattersFolder.folderId) throw new Error("Cannot resolve matters folder for client");

    matterFolderUrl = crm_getOrCreateMatterFolder(matter.MATTER_ID, matter.TITLE || matter.MATTER_ID, mattersFolder.folderId);
    if (matterFolderUrl) {
      crm_setMatterField(matter.MATTER_ID, "FOLDER_URL", matterFolderUrl);
      matter.FOLDER_URL = matterFolderUrl;
    }
  }

  const folderId = extractFolderIdFromUrl_(matter.FOLDER_URL || "") || extractFolderIdFromUrl_(client.FOLDER_URL);
  if (!folderId) throw new Error("Could not resolve drive folder ID for matter/client");

  // PART 5: Debug log — surface missing office profile fields before generation
  var office_ = crm_getOfficeProfile();
  var missingProps = [];
  if (!office_.LAWYERS_RU && !office_.LAWYERS_HE) missingProps.push("LAWYERS_RU / LAWYERS_HE");
  if (!office_.LAW_OFFICE_RU && !office_.LAW_OFFICE_HE) missingProps.push("LAW_OFFICE_RU / LAW_OFFICE_HE");
  if (!office_.OFFICE_NAME && !office_.OFFICE_NAME_HE) missingProps.push("OFFICE_NAME_RU / OFFICE_NAME_HE");
  logInfo_("GENDOC", "Generating Agreement+POA: matterId=" + matterId +
    " clientId=" + client.CLIENT_ID +
    " clientName=" + (client.FULL_NAME || "(empty)") +
    " idNumber=" + (String(client.ID_NUMBER || "").trim() || "(empty)") +
    " authority=" + (matter.AUTHORITY || "(empty)") +
    " folderId=" + folderId +
    (missingProps.length ? " | MISSING_PROPS=[" + missingProps.join(", ") + "]" : ""), {});

  const createdDocs = [];

  // Generate agreement
  const agreementDoc = crm_createGoogleDocFromTemplate(
    `Agreement_${matter.MATTER_ID}_${Date.now()}`,
    crm_buildAgreementText(client, matter),
    folderId
  );

  const agreementPdf = crm_exportDocToPdf(agreementDoc.getId(), folderId, `Agreement_${matter.MATTER_ID}.pdf`);

  const agreementAdd = crm_addDocument({
    clientId: client.CLIENT_ID,
    matterId: matter.MATTER_ID,
    type: "AGREEMENT",
    status: "DRAFT",
    title: `Agreement: ${matter.TITLE || matter.MATTER_ID}`,
    docUrl: agreementDoc.getUrl(),
    pdfUrl: agreementPdf.getUrl(),
    fileId: agreementDoc.getId(),
    createdBy: getActiveUserEmail_() || "system",
    notes: `Generated by CRM sign flow on ${nowIso_()}`,
  });

  createdDocs.push(agreementAdd);

  // Generate POA
  const poaDoc = crm_createGoogleDocFromTemplate(
    `POA_${matter.MATTER_ID}_${Date.now()}`,
    crm_buildPoaText(client, matter),
    folderId
  );

  const poaPdf = crm_exportDocToPdf(poaDoc.getId(), folderId, `POA_${matter.MATTER_ID}.pdf`);

  const poaAdd = crm_addDocument({
    clientId: client.CLIENT_ID,
    matterId: matter.MATTER_ID,
    type: "POA",
    status: "DRAFT",
    title: `POA: ${matter.TITLE || matter.MATTER_ID}`,
    docUrl: poaDoc.getUrl(),
    pdfUrl: poaPdf.getUrl(),
    fileId: poaDoc.getId(),
    createdBy: getActiveUserEmail_() || "system",
    notes: `Generated by CRM sign flow on ${nowIso_()}`,
  });

  createdDocs.push(poaAdd);

  crm_logActivity({
    action: "SIGN_DOCS_GENERATED",
    message: `Generated Agreement + POA for matter ${matterId}`,
    clientId: client.CLIENT_ID,
    matterId: matterId,
    meta: {
      documents: createdDocs.map((d) => ({ docId: d.docId, type: d.type || "" })),
      folderUrl: matter.FOLDER_URL,
    },
  });

  // Onboarding matters must stay at ONBOARDING stage so crm_createSignLinksForMatter
  // continues to use the package-token path (not the per-doc path) when "Create Sign Links"
  // is clicked after doc generation.
  if ((matter.STAGE || "").toUpperCase() !== "ONBOARDING") {
    crm_updateMatterStage(matterId, "DOCUMENTS_GENERATED");
  }

  return {
    ok: true,
    matterId,
    clientId: client.CLIENT_ID,
    folderUrl: matter.FOLDER_URL,
    agreement: { docId: agreementAdd.docId, docUrl: agreementDoc.getUrl(), pdfUrl: agreementPdf.getUrl() },
    poa: { docId: poaAdd.docId, docUrl: poaDoc.getUrl(), pdfUrl: poaPdf.getUrl() },
  };
}

function crm_createGoogleDocFromTemplate(name, textContent, folderId) {
  const doc = DocumentApp.create(name);
  const body = doc.getBody();
  body.clear();

  const lines = String(textContent || "").split("\n");
  lines.forEach(function (line) {
    body.appendParagraph(line);
  });

  doc.saveAndClose();

  const file = DriveApp.getFileById(doc.getId());
  if (folderId) {
    const folder = DriveApp.getFolderById(folderId);
    folder.addFile(file);
    try { DriveApp.getRootFolder().removeFile(file); } catch (e) { /* ignore root removal failures */ }
  }

  return doc;
}

function crm_exportDocToPdf(docId, folderId, fileName) {
  if (!docId) throw new Error("crm_exportDocToPdf: docId is required");
  if (!fileName) fileName = `doc_${Date.now()}.pdf`;

  const blob = DriveApp.getFileById(docId).getBlob().getAs("application/pdf").setName(fileName);
  if (folderId) {
    const folder = DriveApp.getFolderById(folderId);
    return folder.createFile(blob);
  }

  return DriveApp.createFile(blob);
}

function crm_buildAgreementText(client, matter) {
  const data = crm_buildTemplateData(client, matter);
  return crm_templateFill(CRM_TEMPLATE_AGREEMENT, data);
}

function crm_buildPoaText(client, matter) {
  const data = crm_buildTemplateData(client, matter);
  return crm_templateFill(CRM_TEMPLATE_POA, data);
}

function crm_setMatterField(matterId, field, value) {
  if (!matterId || !field) throw new Error("crm_setMatterField: matterId and field are required");

  const fieldTrimmed = String(field).trim();
  const ss = crm_getSpreadsheet_();
  const c = cfg_();
  const sh = ss.getSheetByName(c.SHEETS.MATTERS);
  if (!sh) throw new Error("Matters sheet not found");

  var headers = sh.getLastColumn() > 0 ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
  var idxMatter = headers.indexOf("MATTER_ID");
  var idxField = headers.indexOf(fieldTrimmed);

  // Auto-ensure: if the column is missing, run ensureHeaders_ once and retry
  if (idxField === -1) {
    ensureHeaders_(sh, c.HEADERS.MATTERS);
    headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    idxMatter = headers.indexOf("MATTER_ID");
    idxField = headers.indexOf(fieldTrimmed);
    logInfo_("SETUP", "crm_setMatterField: auto-ensured MATTERS headers for field: " + fieldTrimmed, {});
  }

  if (idxMatter === -1 || idxField === -1) {
    throw new Error("crm_setMatterField: field not found: \"" + fieldTrimmed + "\". Sheet headers: [" + headers.join(", ") + "]");
  }

  const values = sh.getRange(2, idxMatter + 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(matterId)) {
      sh.getRange(i + 2, idxField + 1).setValue(value);
      return true;
    }
  }

  return false;
}

function crm_createSignLinkForDocument(docId) {
  if (!docId) throw new Error("crm_createSignLinkForDocument: docId is required");

  const lock = LockService.getScriptLock();
  const locked = lock.tryLock(10000);
  if (!locked) throw new Error("crm_createSignLinkForDocument: could not acquire lock");

  try {
    const doc = crm_getDocumentById(docId);
    if (!doc) throw new Error("Document not found: " + docId);

    // Reuse existing active token for same document to prevent token spam
    const existing = crm_findActiveSignTokenForDoc(docId);
    if (existing) {
      const link = crm_getWebAppUrl_("sign") + "&token=" + encodeURIComponent(existing.token);
      return {
        ok: true,
        docId,
        docType: doc.TYPE || "DOCUMENT",
        docTitle: doc.TITLE || "Document",
        docStatus: "PENDING",
        signUrl: link,
        token: existing.token,
        reused: true,
      };
    }

  const token = Utilities.getUuid() + "-" + Utilities.getUuid();
  const expiresAt = addHours_(new Date(), CRM_SIGN_TTL_HOURS).toISOString();
  const url = crm_getWebAppUrl_("sign") + "&token=" + encodeURIComponent(token);

  const leadId = doc.CLIENT_ID ? crm_getLeadIdByClientId(doc.CLIENT_ID) : null;

  const entry = {
    token,
    docId: docId,
    matterId: doc.MATTER_ID || "",
    clientId: doc.CLIENT_ID || "",
    leadId: leadId || "",
    status: CRM_SIGN_TOKEN_STATUS_ACTIVE,
    createdAt: new Date().toISOString(),
    expiresAt: expiresAt,
  };

  crm_setSignTokenEntry(token, entry);

  crm_updateDocumentStatus(docId, "PENDING", `Sign link created: ${url}`);

  crm_logActivity({
    action: "SIGN_LINK_CREATED",
    message: `Sign link created for document ${docId}`,
    clientId: doc.CLIENT_ID,
    matterId: doc.MATTER_ID,
    meta: { docId, token, signUrl: url, status: entry.status },
  });

  return {
    ok: true,
    docId,
    docType: doc.TYPE || "DOCUMENT",
    docTitle: doc.TITLE || "Document",
    docStatus: "PENDING",
    signUrl: url,
    token,
  };
  } finally {
    lock.releaseLock();
  }
}

function crm_createSignLinksForMatter(matterId) {
  if (!matterId) throw new Error("crm_createSignLinksForMatter: matterId is required");

  // For ONBOARDING matters: one package link covers Agreement + POA
  const matter = crm_getMatterById(matterId);
  if (matter && (matter.STAGE || "").toUpperCase() === "ONBOARDING") {
    try {
      const pkg = crm_createOnboardingSignPackage_(matterId);
      return { ok: true, matterId, links: [pkg] };
    } catch (e) {
      return { ok: true, matterId, links: [{ ok: false, error: e.message }] };
    }
  }

  // Non-onboarding: per-document links (unchanged)
  const docs = crm_listDocumentsByMatterId(matterId, { limit: 100 });
  const links = docs
    .filter((d) => ["AGREEMENT", "POA"].includes((d.TYPE || "").toUpperCase()))
    .map((d) => {
      try {
        const res = crm_createSignLinkForDocument(d.DOC_ID);
        return res;
      } catch (e) {
        return { ok: false, error: e.message, docId: d.DOC_ID };
      }
    });

  return { ok: true, matterId, links };
}

function crm_getSignInfo(token) {
  if (!token) throw new Error("crm_getSignInfo: token is required");

  const entry = crm_getSignTokenEntry(token);
  if (!entry || !entry.docId) throw new Error("Invalid token");

  if (entry.status === CRM_SIGN_TOKEN_STATUS_USED) {
    // If signing was fully completed, return a clean "already signed" info response
    // so the sign page can reach the terminal "already signed" state instead of an error.
    if (entry.kind === "ONBOARDING_PACKAGE" && Array.isArray(entry.docIds)) {
      const allSigned = entry.docIds.every(function(id) {
        const d = crm_getDocumentById(id);
        return d && (d.STATUS || "").toUpperCase() === "SIGNED";
      });
      if (allSigned) {
        const m0 = entry.matterId ? crm_getMatterById(entry.matterId) : null;
        const cl0 = entry.clientId ? crm_findClientById(entry.clientId) : null;
        return {
          ok: true, token,
          doc: { TITLE: "Agreement + POA (Onboarding)", STATUS: "SIGNED", MATTER_ID: entry.matterId, CLIENT_ID: entry.clientId },
          matter: m0, client: cl0, leadId: entry.leadId || null,
          signStatus: "SIGNED",
          signUrl: crm_getWebAppUrl_("sign") + "&token=" + encodeURIComponent(token),
        };
      }
    } else {
      // Regular single-doc token: if the doc is signed, return clean response
      const docCheck = crm_getDocumentById(entry.docId);
      if (docCheck && (docCheck.STATUS || "").toUpperCase() === "SIGNED") {
        const m1 = docCheck.MATTER_ID ? crm_getMatterById(docCheck.MATTER_ID) : null;
        const cl1 = docCheck.CLIENT_ID ? crm_findClientById(docCheck.CLIENT_ID) : null;
        return {
          ok: true, token,
          doc: docCheck, matter: m1, client: cl1, leadId: entry.leadId || null,
          signStatus: "SIGNED",
          signUrl: crm_getWebAppUrl_("sign") + "&token=" + encodeURIComponent(token),
        };
      }
    }
    // Token used but signing was not fully completed — throw so sign page failure handler terminates
    throw new Error("Token already used. Please request a new signing link.");
  }

  if (crm_isTokenExpired(entry)) {
    entry.status = CRM_SIGN_TOKEN_STATUS_EXPIRED;
    crm_setSignTokenEntry(token, entry);
    throw new Error("Token expired. Please request a new signing link.");
  }

  if (entry.status !== CRM_SIGN_TOKEN_STATUS_ACTIVE) {
    throw new Error("Token is not active. Please request a new signing link.");
  }

  const doc = crm_getDocumentById(entry.docId);
  if (!doc) throw new Error("Document not found for token");

  const matter = doc.MATTER_ID ? crm_getMatterById(doc.MATTER_ID) : null;
  // Provisional onboarding: CLIENT_ID may be a leadId — show lead data on sign page
  var client = null;
  if (doc.CLIENT_ID && !String(doc.CLIENT_ID).startsWith("LEAD_")) {
    client = crm_findClientById(doc.CLIENT_ID);
  } else if (entry.leadId || (doc.CLIENT_ID && String(doc.CLIENT_ID).startsWith("LEAD_"))) {
    const leadForSign_ = crm_getLeadById(entry.leadId || doc.CLIENT_ID);
    if (leadForSign_) {
      client = { CLIENT_ID: doc.CLIENT_ID, FULL_NAME: leadForSign_.FULL_NAME, PHONE: leadForSign_.PHONE, EMAIL: leadForSign_.EMAIL };
    }
  }

  // For onboarding package: signStatus = SIGNED only when all docs in the package are signed
  var displayDoc = doc;
  var signStatus = (doc.STATUS || "").toUpperCase();
  if (entry.kind === "ONBOARDING_PACKAGE" && Array.isArray(entry.docIds)) {
    const allSigned = entry.docIds.every(function(id) {
      const d = crm_getDocumentById(id);
      return d && (d.STATUS || "").toUpperCase() === "SIGNED";
    });
    signStatus = allSigned ? "SIGNED" : "PENDING";
    displayDoc = { TITLE: "Agreement + POA (Onboarding)", STATUS: signStatus, MATTER_ID: doc.MATTER_ID, CLIENT_ID: doc.CLIENT_ID };
  }

  return {
    ok: true,
    token,
    doc: displayDoc,
    matter,
    client,
    leadId: entry.leadId || null,
    signStatus: signStatus,
    signUrl: crm_getWebAppUrl_("sign") + "&token=" + encodeURIComponent(token),
  };
}

// Returns the first SIGNED_ output document of the given type in a matter.
// Used to detect existing signed output before creating duplicates.
function crm_findSignedOutputForDoc_(matterId, signedType) {
  if (!matterId || !signedType) return null;
  const upper = signedType.toUpperCase();
  const docs = crm_listDocumentsByMatterId(matterId, { limit: 100 });
  return docs.find(function(d) {
    return (d.TYPE || "").toUpperCase() === upper && (d.STATUS || "").toUpperCase() === "SIGNED";
  }) || null;
}

function crm_submitSignature(token, dataUrlPng) {
  if (!token) throw new Error("crm_submitSignature: token is required");
  if (!dataUrlPng || !String(dataUrlPng).startsWith("data:image/png;base64,")) throw new Error("Bad signature payload");

  // Serialize concurrent submits — prevents duplicate PNG/PDF if two requests race in
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) {
    return { ok: false, code: "processing", message: "Signature is already being processed. Please wait.", error: "Signature is already being processed. Please wait." };
  }

  try {
    const entry = crm_getSignTokenEntry(token);
    if (!entry || !entry.docId) throw new Error("Invalid token");

    // Token already consumed — distinguish: signed OK vs. signing incomplete
    if (entry.status === CRM_SIGN_TOKEN_STATUS_USED) {
      // For onboarding package: all docIds must be signed for already_signed
      if (entry.kind === "ONBOARDING_PACKAGE" && Array.isArray(entry.docIds)) {
        const allSigned = entry.docIds.every(function(id) {
          const d = crm_getDocumentById(id);
          return d && (d.STATUS || "").toUpperCase() === "SIGNED";
        });
        if (allSigned) {
          return { ok: true, already: true, code: "already_signed", message: "Documents already signed" };
        }
        return { ok: false, code: "token_used", message: "Token already used. Please request a new signing link.", error: "Token already used. Please request a new signing link." };
      }
      // Single-doc: check primary docId
      const docCheck = crm_getDocumentById(entry.docId);
      if (docCheck && (docCheck.STATUS || "").toUpperCase() === "SIGNED") {
        const signedType = "SIGNED_" + (docCheck.TYPE || "").toUpperCase();
        const prev = crm_findSignedOutputForDoc_(entry.matterId, signedType);
        return { ok: true, already: true, code: "already_signed", docId: entry.docId, signedPdfUrl: prev ? prev.PDF_URL || "" : "", message: "Document already signed" };
      }
      // Token used but doc not properly signed — caller must request a new link
      return { ok: false, code: "token_used", message: "Token already used. Please request a new signing link.", error: "Token already used. Please request a new signing link." };
    }

    if (crm_isTokenExpired(entry)) {
      crm_expireSignToken(token);
      throw new Error("Token expired. Please request a new signing link.");
    }

    if (entry.status !== CRM_SIGN_TOKEN_STATUS_ACTIVE) {
      throw new Error("Token is not active. Please request a new signing link.");
    }

    // ─── ONBOARDING PACKAGE FLOW ────────────────────────────────────────────
    if (entry.kind === "ONBOARDING_PACKAGE" && Array.isArray(entry.docIds)) {
      const pkgMatter = crm_getMatterById(entry.matterId);
      const pkgClient = crm_findClientById(entry.clientId);
      const pkgFolderId = pkgMatter && pkgMatter.FOLDER_URL ? extractFolderIdFromUrl_(pkgMatter.FOLDER_URL)
        : pkgClient && pkgClient.FOLDER_URL ? extractFolderIdFromUrl_(pkgClient.FOLDER_URL) : null;
      if (!pkgFolderId) throw new Error("Could not resolve folder to save signature");

      // Decode PNG blob in memory — Drive file is saved AFTER all docs succeed
      // to avoid orphan signature files accumulating in Drive on partial failure/retry.
      const pkgB64 = String(dataUrlPng).split(",")[1];
      const pkgBytes = Utilities.base64Decode(pkgB64);
      const pkgSigBlob = Utilities.newBlob(pkgBytes, "image/png", "signature_onboarding_" + token + ".png");
      var pkgLastPdfUrl = "";

      for (var i = 0; i < entry.docIds.length; i++) {
        var pkgDocId = entry.docIds[i];
        var pkgDoc = crm_getDocumentById(pkgDocId);
        if (!pkgDoc) throw new Error("Onboarding package document not found: " + pkgDocId);
        var pkgSignedType = "SIGNED_" + (pkgDoc.TYPE || "").toUpperCase();

        // Idempotency: skip if already signed
        if ((pkgDoc.STATUS || "").toUpperCase() === "SIGNED") continue;

        // Idempotency: skip if signed output already exists in sheet
        if (crm_findSignedOutputForDoc_(entry.matterId, pkgSignedType)) {
          crm_updateDocumentStatus(pkgDocId, "SIGNED", "Recovered: signed output already existed");
          continue;
        }

        var pkgPdfFile = crm_makeSignedPdfFromDoc_(
          pkgDoc.DOC_URL || "", pkgFolderId,
          "SIGNED_" + (pkgDoc.TYPE || "DOC") + "_" + pkgDocId + ".pdf",
          pkgSigBlob, token + "_" + pkgDocId
        );
        crm_updateDocumentStatus(pkgDocId, "SIGNED", "Signed via onboarding package token " + token);
        crm_addDocument({
          clientId: entry.clientId, matterId: entry.matterId,
          type: pkgSignedType, status: "SIGNED",
          title: "Signed " + (pkgDoc.TYPE || "Document") + ": " + (pkgDoc.TITLE || ""),
          docUrl: "", pdfUrl: pkgPdfFile.getUrl(), fileId: pkgPdfFile.getId(),
          createdBy: getActiveUserEmail_() || "system",
          notes: "Onboarding package token: " + token + " at " + nowIso_(),
        });
        crm_logActivity({
          action: "DOCUMENT_SIGNED",
          message: "Document signed via onboarding package token " + token,
          clientId: entry.clientId, matterId: entry.matterId,
          meta: { docId: pkgDocId, signedPdfUrl: pkgPdfFile.getUrl() },
        });
        pkgLastPdfUrl = pkgPdfFile.getUrl();
      }

      // Save signature PNG to Drive only after all docs succeed — no orphan files on partial failure
      const pkgSigFile = DriveApp.getFolderById(pkgFolderId).createFile(pkgSigBlob);

      crm_logActivity({
        action: "ONBOARDING_PACKAGE_SIGNED",
        message: "Onboarding package (Agreement + POA) fully signed via token " + token,
        clientId: entry.clientId,
        matterId: entry.matterId,
        meta: { token, signatureUrl: pkgSigFile.getUrl(), signedPdfUrl: pkgLastPdfUrl },
      });

      crm_markSignTokenUsed(token);
      var convertedClientId = null;
      if (entry.matterId && crm_hasBothSignedOnboardingDocs_(entry.matterId)) {
        if (entry.leadId) {
          // Signing-first: create real client now, reattach matter + docs
          const finalClientId = crm_finalizeOnboardingConversion_(entry.leadId, entry.matterId);
          convertedClientId = finalClientId || null;
          crm_triggerNextModuleReadiness_(entry.matterId, finalClientId || entry.clientId);
        } else {
          // Legacy token without leadId: fall back to status-only conversion
          crm_safelyConvertLeadForClient(entry.clientId);
          convertedClientId = (!crm_isProvisionalClientId_(entry.clientId) ? entry.clientId : null);
          crm_triggerNextModuleReadiness_(entry.matterId, entry.clientId);
        }
      }
      crm_logActivity({
        action: "SIGNATURE_SUCCESS",
        message: "Onboarding package signed and conversion complete. clientId: " + (convertedClientId || entry.clientId),
        clientId: convertedClientId || entry.clientId,
        matterId: entry.matterId,
        meta: { token, convertedClientId },
      });
      return { ok: true, code: "success", matterId: entry.matterId || null, clientId: convertedClientId, signedPdfUrl: pkgLastPdfUrl, signatureUrl: pkgSigFile.getUrl() };
    }
    // ─── END ONBOARDING PACKAGE FLOW ────────────────────────────────────────

    const doc = crm_getDocumentById(entry.docId);
    if (!doc) throw new Error("Document not found: " + entry.docId);

    const signedType = "SIGNED_" + (doc.TYPE || "").toUpperCase();

    // Idempotency: doc already SIGNED (completed by an earlier request that won the lock)
    if ((doc.STATUS || "").toUpperCase() === "SIGNED") {
      crm_markSignTokenUsed(token);
      const prev = crm_findSignedOutputForDoc_(doc.MATTER_ID, signedType);
      return { ok: true, already: true, code: "already_signed", docId: entry.docId, signedPdfUrl: prev ? prev.PDF_URL || "" : "", message: "Document already signed" };
    }

    const matter = doc.MATTER_ID ? crm_getMatterById(doc.MATTER_ID) : null;
    const client = doc.CLIENT_ID ? crm_findClientById(doc.CLIENT_ID) : null;

    const folderId = matter?.FOLDER_URL ? extractFolderIdFromUrl_(matter.FOLDER_URL)
      : client?.FOLDER_URL ? extractFolderIdFromUrl_(client.FOLDER_URL) : null;
    if (!folderId) throw new Error("Could not resolve folder to save signature");

    // Guard: signed output already exists for this doc type — reuse, do not duplicate
    const existingSignedDoc = crm_findSignedOutputForDoc_(doc.MATTER_ID, signedType);
    if (existingSignedDoc) {
      crm_updateDocumentStatus(entry.docId, "SIGNED", "Recovered: signed output already existed");
      crm_markSignTokenUsed(token);
      return { ok: true, already: true, code: "already_signed", docId: entry.docId, signedPdfUrl: existingSignedDoc.PDF_URL || "", message: "Document already signed" };
    }

    const b64 = String(dataUrlPng).split(",")[1];
    const bytes = Utilities.base64Decode(b64);
    const sigBlob = Utilities.newBlob(bytes, "image/png", `signature_${entry.docId}.png`);
    const sigFile = DriveApp.getFolderById(folderId).createFile(sigBlob);

    // create signed PDF version with marker replacement
    const signedPdfFile = crm_makeSignedPdfFromDoc_(doc.DOC_URL || "", folderId, `SIGNED_${doc.TYPE || 'DOC'}_${doc.DOC_ID}.pdf`, sigBlob, token);

    // update original document status
    crm_updateDocumentStatus(entry.docId, "SIGNED", `Signed by token ${token}`);

    crm_markSignTokenUsed(token);

    crm_addDocument({
      clientId: doc.CLIENT_ID,
      matterId: doc.MATTER_ID,
      type: signedType,
      status: "SIGNED",
      title: `Signed ${doc.TYPE || "Document"}: ${doc.TITLE || ""}`,
      docUrl: "",
      pdfUrl: signedPdfFile.getUrl(),
      fileId: signedPdfFile.getId(),
      createdBy: getActiveUserEmail_() || "system",
      notes: `Signature file ${sigFile.getUrl()} stored at ${nowIso_()}`,
    });

    crm_logActivity({
      action: "DOCUMENT_SIGNED",
      message: `Document signed via token ${token}`,
      clientId: doc.CLIENT_ID,
      matterId: doc.MATTER_ID,
      meta: { docId: entry.docId, signedPdfUrl: signedPdfFile.getUrl(), signatureImageUrl: sigFile.getUrl() },
    });

    // Convert lead only after BOTH AGREEMENT and POA are signed for this matter
    if (doc.MATTER_ID && crm_hasBothSignedOnboardingDocs_(doc.MATTER_ID)) {
      if (doc.CLIENT_ID && String(doc.CLIENT_ID).startsWith("LEAD_")) {
        // Provisional: finalize by creating real client from lead data
        const finalClientId = crm_finalizeOnboardingConversion_(doc.CLIENT_ID, doc.MATTER_ID);
        crm_triggerNextModuleReadiness_(doc.MATTER_ID, finalClientId || doc.CLIENT_ID);
      } else {
        crm_safelyConvertLeadForClient(doc.CLIENT_ID);
        crm_triggerNextModuleReadiness_(doc.MATTER_ID, doc.CLIENT_ID);
      }
    }

    // optional: expire token
    PropertiesService.getScriptProperties().setProperty(CRM_SIGN_TOKEN_PREFIX + token + "_expires", new Date().toISOString());

    return {
      ok: true,
      code: "success",
      docId: entry.docId,
      signedPdfUrl: signedPdfFile.getUrl(),
      signatureUrl: sigFile.getUrl(),
    };

  } finally {
    lock.releaseLock();
  }
}

function crm_hasBothSignedOnboardingDocs_(matterId) {
  if (!matterId) return false;
  const docs = crm_listDocumentsByMatterId(matterId, { limit: 100 });
  const isSigned = function(type) {
    return docs.some(function(d) {
      return (d.TYPE || "").toUpperCase() === type && (d.STATUS || "").toUpperCase() === "SIGNED";
    });
  };
  return isSigned("AGREEMENT") && isSigned("POA");
}

function crm_makeSignedPdfFromDoc_(docUrl, folderId, pdfName, sigBlob, token) {
  if (!docUrl) throw new Error("crm_makeSignedPdfFromDoc_: docUrl is required");
  const docId = extractIdFromUrl_(docUrl);
  if (!docId) throw new Error("crm_makeSignedPdfFromDoc_: could not resolve doc ID");

  const tempFile = DriveApp.getFileById(docId).makeCopy(`TEMP_SIGN_${token}_${Date.now()}`, DriveApp.getFolderById(folderId));
  const tempDocId = tempFile.getId();
  const doc = DocumentApp.openById(tempDocId);

  insertSignatureByMarker_(doc, sigBlob, CRM_SIGN_MARKER);
  doc.saveAndClose();

  const pdfBlob = DriveApp.getFileById(tempDocId).getBlob().getAs("application/pdf").setName(pdfName);
  const pdfFile = DriveApp.getFolderById(folderId).createFile(pdfBlob);

  tempFile.setTrashed(true);

  return pdfFile;
}

function insertSignatureByMarker_(doc, sigBlob, markerText) {
  const body = doc.getBody();
  const found = body.findText(escapeForRegex_(markerText));
  if (!found) throw new Error("Signature marker not found as part of signed PDF generation");

  const element = found.getElement();
  const start = found.getStartOffset();
  const end = found.getEndOffsetInclusive();

  const text = element.asText();
  text.deleteText(start, end);

  const parent = getParentParagraph_(element);
  if (!parent) throw new Error("Cannot find parent paragraph for marker");

  const img = parent.insertInlineImage(parent.getChildIndex(element) + 1, sigBlob);
  img.setWidth(Math.min(420, img.getWidth()));
  img.setHeight(Math.min(160, img.getHeight()));
  parent.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  return true;
}

function getParentParagraph_(element) {
  let el = element;
  while (el) {
    const t = el.getType();
    if (t === DocumentApp.ElementType.PARAGRAPH) return el.asParagraph();
    if (t === DocumentApp.ElementType.LIST_ITEM) return el.asListItem();
    el = el.getParent();
  }
  return null;
}

function escapeForRegex_(text) {
  return String(text).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function addHours_(dt, hours) {
  const d = (dt instanceof Date) ? new Date(dt.getTime()) : new Date(dt);
  d.setHours(d.getHours() + hours);
  return d;
}

// Scan ScriptProperties for an active ONBOARDING_PACKAGE token for the given matter.
function crm_findActiveOnboardingPackageForMatter_(matterId) {
  if (!matterId) return null;
  const all = PropertiesService.getScriptProperties().getProperties();
  const prefix = CRM_SIGN_TOKEN_PREFIX;
  for (const key in all) {
    if (!key.startsWith(prefix)) continue;
    if (key.endsWith("_expires")) continue;
    let entry;
    try { entry = JSON.parse(all[key]); } catch(e) { continue; }
    if (!entry || entry.kind !== "ONBOARDING_PACKAGE") continue;
    if (entry.matterId !== matterId) continue;
    if (entry.status !== CRM_SIGN_TOKEN_STATUS_ACTIVE) continue;
    if (crm_isTokenExpired(entry)) {
      entry.status = CRM_SIGN_TOKEN_STATUS_EXPIRED;
      crm_setSignTokenEntry(entry.token, entry);
      continue;
    }
    return entry;
  }
  return null;
}

// Create (or reuse) ONE signing link that covers both AGREEMENT and POA for an onboarding matter.
// Called only by crm_createSignLinksForMatter when matter.STAGE === "ONBOARDING".
function crm_createOnboardingSignPackage_(matterId) {
  const docs = crm_listDocumentsByMatterId(matterId, { limit: 100 });
  const agDoc  = docs.find(function(d) { return (d.TYPE || "").toUpperCase() === "AGREEMENT"; });
  const poaDoc = docs.find(function(d) { return (d.TYPE || "").toUpperCase() === "POA"; });

  if (!agDoc || !poaDoc) {
    throw new Error("Generate Agreement + POA first before creating the onboarding sign link.");
  }

  // Block: both docs already signed — no new link needed
  const agSigned  = (agDoc.STATUS  || "").toUpperCase() === "SIGNED";
  const poaSigned = (poaDoc.STATUS || "").toUpperCase() === "SIGNED";
  if (agSigned && poaSigned) {
    crm_logActivity({
      action: "DUPLICATE_SIGN_LINK_BLOCKED",
      message: "New onboarding sign link blocked — both docs already signed for matter " + matterId,
      matterId: matterId,
      clientId: agDoc.CLIENT_ID || "",
      meta: {},
    });
    return {
      ok: false,
      code: "already_signed",
      message: "Both Agreement and POA are already signed for this matter.",
    };
  }

  // Reuse existing active package token for this matter (idempotent)
  const existing = crm_findActiveOnboardingPackageForMatter_(matterId);
  if (existing) {
    const link = crm_getWebAppUrl_("sign") + "&token=" + encodeURIComponent(existing.token);
    crm_logActivity({
      action: "ONBOARDING_PACKAGE_LINK_REUSED",
      message: "Onboarding sign package link resent for matter " + matterId,
      clientId: agDoc.CLIENT_ID,
      matterId: matterId,
      meta: { token: existing.token, signUrl: link },
    });
    return {
      ok: true,
      docType: "ONBOARDING_PACKAGE",
      docTitle: "Agreement + POA (Onboarding)",
      docStatus: "PENDING",
      signUrl: link,
      token: existing.token,
      whatsappUrl: crm_buildWhatsAppUrl_("Please sign your onboarding documents (Agreement + POA): " + link),
      reused: true,
    };
  }

  const token = Utilities.getUuid() + "-" + Utilities.getUuid();
  const expiresAt = addHours_(new Date(), CRM_SIGN_TTL_HOURS).toISOString();
  const url = crm_getWebAppUrl_("sign") + "&token=" + encodeURIComponent(token);
  // When using provisional design, agDoc.CLIENT_ID IS the leadId (LEAD_xxx).
  // For real clients, derive leadId by looking up who is linked to that client.
  const provisionalCid_ = agDoc.CLIENT_ID || "";
  const leadId = provisionalCid_.startsWith("LEAD_")
    ? provisionalCid_
    : (provisionalCid_ ? crm_getLeadIdByClientId(provisionalCid_) : null);

  const entry = {
    token,
    kind: "ONBOARDING_PACKAGE",
    docId: agDoc.DOC_ID,                         // compatibility anchor for crm_getSignInfo
    docIds: [agDoc.DOC_ID, poaDoc.DOC_ID],
    matterId: agDoc.MATTER_ID || "",
    clientId: agDoc.CLIENT_ID || "",
    leadId: leadId || "",
    status: CRM_SIGN_TOKEN_STATUS_ACTIVE,
    createdAt: new Date().toISOString(),
    expiresAt: expiresAt,
  };

  crm_setSignTokenEntry(token, entry);

  // Mark both source docs PENDING
  crm_updateDocumentStatus(agDoc.DOC_ID,  "PENDING", "Onboarding sign package created: " + url);
  crm_updateDocumentStatus(poaDoc.DOC_ID, "PENDING", "Onboarding sign package created: " + url);

  crm_logActivity({
    action: "ONBOARDING_PACKAGE_LINK_CREATED",
    message: "Onboarding sign package (Agreement + POA) created for matter " + matterId,
    clientId: agDoc.CLIENT_ID,
    matterId: matterId,
    meta: { docIds: entry.docIds, token, signUrl: url },
  });

  return {
    ok: true,
    docType: "ONBOARDING_PACKAGE",
    docTitle: "Agreement + POA (Onboarding)",
    docStatus: "PENDING",
    signUrl: url,
    token,
    whatsappUrl: crm_buildWhatsAppUrl_("Please sign your onboarding documents (Agreement + POA): " + url),
    reused: false,
  };
}

// Builds a WhatsApp deep link with the given message text.
// Usage: open the returned URL in a browser tab to pre-fill the WhatsApp share dialog.
function crm_buildWhatsAppUrl_(messageText) {
  return "https://wa.me/?text=" + encodeURIComponent(String(messageText || ""));
}

// Returns true when an id looks like a provisional lead-based CLIENT_ID (LEAD_xxx).
// Any value that starts with "LEAD_" was assigned as a provisional stand-in
// for a real client and must never be treated as a final client reference.
function crm_isProvisionalClientId_(id) {
  return Boolean(id) && String(id).startsWith("LEAD_");
}

// Creates a real client from lead data after onboarding signing completes.
// Updates lead CLIENT_ID + STATUS = CONVERTED.
// Reattaches onboarding matter + all its documents to the new real client.
// Idempotent: if lead is already CONVERTED with a real client, returns existing clientId.
function crm_finalizeOnboardingConversion_(leadId, matterId) {
  if (!leadId || !matterId) return null;
  const lead = crm_getLeadById(leadId);
  if (!lead) return null;

  // Idempotency: if lead already has a real CLIENT_ID, skip creation regardless of STATUS.
  // STATUS may still be stale if a previous run was interrupted between writing CLIENT_ID
  // and writing STATUS=CONVERTED — using CLIENT_ID alone avoids a duplicate client.
  const existingClientId = String(lead.CLIENT_ID || "").trim();
  if (existingClientId && !crm_isProvisionalClientId_(existingClientId)) {
    // Repair: matter.CLIENT_ID may still be provisional if conversion was interrupted after
    // crm_addClient but before crm_setMatterField (or crm_reattachMatterDocsToClient_).
    const matterForRepair = crm_getMatterById(matterId);
    if (matterForRepair && crm_isProvisionalClientId_(matterForRepair.CLIENT_ID)) {
      crm_setMatterField(matterId, "CLIENT_ID", existingClientId);
      crm_reattachMatterDocsToClient_(matterId, leadId, existingClientId);
      crm_logActivity({
        action: "ONBOARDING_CONVERSION_REPAIRED",
        message: "Completed partial conversion: matter/docs reattached to existing client " + existingClientId,
        clientId: existingClientId,
        matterId: matterId,
        meta: { leadId: leadId, existingClientId: existingClientId, matterId: matterId },
      });
    }
    // Repair STATUS if still not CONVERTED
    if (String(lead.STATUS || "").toUpperCase() !== "CONVERTED" && lead.__row) {
      const c_ = cfg_();
      const sh_ = crm_getSpreadsheet_().getSheetByName(c_.SHEETS.LEADS);
      if (sh_) {
        const headers_ = sh_.getRange(1, 1, 1, sh_.getLastColumn()).getValues()[0];
        const iStatus_ = headers_.indexOf("STATUS");
        if (iStatus_ >= 0) sh_.getRange(lead.__row, iStatus_ + 1).setValue("CONVERTED");
      }
    }
    return existingClientId;
  }

  // Create real client from lead data
  const clientRes = crm_addClient({
    fullName: String(lead.FULL_NAME || "").trim(),
    phone: String(lead.PHONE || "").trim(),
    email: String(lead.EMAIL || "").trim(),
    idType: String(lead.ID_TYPE || "").trim(),
    idNumber: String(lead.ID_NUMBER || "").trim(),
    address: String(lead.ADDRESS || "").trim(),
    source: "ONBOARDING_SIGNED",
    status: "NEW",
    owner: String(lead.ASSIGNED_TO || getActiveUserEmail_() || "").trim(),
  });
  const realClientId = clientRes.clientId;
  if (!realClientId) return null;

  // Update lead row: CLIENT_ID + STATUS = CONVERTED
  const c = cfg_();
  const sh = crm_getSpreadsheet_().getSheetByName(c.SHEETS.LEADS);
  if (sh) {
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const iClient = headers.indexOf("CLIENT_ID");
    const iStatus = headers.indexOf("STATUS");
    const iUpdated = headers.indexOf("UPDATED_AT");
    const now = nowIso_();
    if (iClient >= 0) sh.getRange(lead.__row, iClient + 1).setValue(realClientId);
    if (iStatus >= 0) sh.getRange(lead.__row, iStatus + 1).setValue("CONVERTED");
    if (iUpdated >= 0) sh.getRange(lead.__row, iUpdated + 1).setValue(now);
  }

  // Reattach matter to real client
  crm_setMatterField(matterId, "CLIENT_ID", realClientId);

  // Reattach all documents for this matter from provisional leadId to real client
  crm_reattachMatterDocsToClient_(matterId, leadId, realClientId);

  crm_logActivity({
    action: "ONBOARDING_CONVERSION_COMPLETE",
    message: "Lead " + leadId + " converted to client " + realClientId + " after onboarding signing",
    clientId: realClientId,
    matterId: matterId,
    meta: { leadId, realClientId, matterId },
  });

  return realClientId;
}

// Updates CLIENT_ID on every document in a matter where CLIENT_ID matches fromClientId.
// Used to migrate provisional (lead-based) doc records to the real client after onboarding conversion.
function crm_reattachMatterDocsToClient_(matterId, fromClientId, toClientId) {
  if (!matterId || !fromClientId || !toClientId) return;
  const c = cfg_();
  const sh = crm_getSpreadsheet_().getSheetByName(c.SHEETS.DOCUMENTS);
  if (!sh || sh.getLastRow() < 2) return;
  const values = sh.getDataRange().getValues();
  const header = values[0];
  const iMatter = header.indexOf("MATTER_ID");
  const iClient = header.indexOf("CLIENT_ID");
  if (iMatter < 0 || iClient < 0) return;
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][iMatter] || "") === String(matterId) &&
        String(values[r][iClient] || "") === String(fromClientId)) {
      sh.getRange(r + 1, iClient + 1).setValue(toClientId);
    }
  }
}

// Maps a matter category string to a next-module code.
// Returns null if no module is mapped for this category.
function crm_resolveNextModule_(category) {
  var cat = String(category || "").toUpperCase().trim();
  if (cat === "WORK_ACCIDENT") return "BL_211";
  if (cat === "LABOR" || cat === "WORK" || cat === "EMPLOYMENT") return "LABOR";
  return null;
}

// Fires a NEXT_MODULE_READY activity and creates a follow-up task.
// Called only after onboarding completion (both docs signed + conversion confirmed).
// Best-effort: task failure never propagates.
function crm_triggerNextModuleReadiness_(matterId, clientId) {
  if (!matterId || !clientId) return;
  var matter = crm_getMatterById(matterId);
  var category = matter ? String(matter.CATEGORY || "").toUpperCase().trim() : "";
  var module = crm_resolveNextModule_(category);
  if (!module) return;

  // Dedupe: exit if NEXT_MODULE_READY for this matter+module was already emitted
  try {
    var existing = crm_listActivities({ matterId: matterId, limit: 50 });
    var alreadyFired = existing.some(function(a) {
      if (String(a.ACTION || "") !== "NEXT_MODULE_READY") return false;
      try {
        var m = typeof a.META_JSON === "string" ? JSON.parse(a.META_JSON) : (a.META_JSON || {});
        return String(m.module || "") === module;
      } catch (e) {
        return false;
      }
    });
    if (alreadyFired) return;
  } catch (e) {
    // If the dedupe check itself fails, proceed rather than silently suppressing
  }

  crm_logActivity({
    action: "NEXT_MODULE_READY",
    message: "Next module ready: " + module + " (category: " + category + ")",
    clientId: clientId,
    matterId: matterId,
    meta: { module: module, reason: category, matterId: matterId, clientId: clientId },
  });

  try {
    crm_createTask({
      clientId: clientId,
      matterId: matterId,
      type: "MODULE_LAUNCH",
      title: "Launch " + module + " module",
      priority: "HIGH",
      generatedBy: "NEXT_MODULE_READY",
    });
  } catch (e) {
    // Task creation is best-effort — never block the signing response
  }
}
