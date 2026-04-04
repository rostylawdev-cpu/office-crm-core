/** crmConsistencyAudit.gs
 *
 * Lightweight consistency audit + safe repair for known CRM data drift.
 * Two entry points:
 *   crm_runConsistencyAudit()   — audit only, no writes, logs findings
 *   crm_runConsistencyRepair()  — audit then apply deterministic safe repairs
 *
 * Both can also be called programmatically; they return a structured result object.
 */

// ─── Sheet row loader ──────────────────────────────────────────────────────

/**
 * Loads all data rows from a CRM sheet as plain objects keyed by header name.
 * sheetKey must match a key in cfg_().SHEETS (e.g. "LEADS", "MATTERS", "DOCUMENTS", "ACTIVITIES").
 * limitFromEnd: if > 0, only loads the last N rows (useful for large Activities sheets).
 */
function crm_auditLoadRows_(sheetKey, limitFromEnd) {
  var c = cfg_();
  var sheetName = c.SHEETS[sheetKey];
  if (!sheetName) return [];
  var sh = crm_getSpreadsheet_().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  var values = sh.getDataRange().getValues();
  var header = values[0];
  var dataStart = 1;
  if (limitFromEnd && limitFromEnd > 0 && values.length - 1 > limitFromEnd) {
    dataStart = values.length - limitFromEnd;
  }
  var out = [];
  for (var r = dataStart; r < values.length; r++) {
    var obj = {};
    for (var col = 0; col < header.length; col++) {
      if (header[col]) obj[String(header[col])] = (values[r][col] !== undefined && values[r][col] !== null) ? values[r][col] : "";
    }
    out.push(obj);
  }
  return out;
}

// ─── Stale token scanner ───────────────────────────────────────────────────

/**
 * Scans ScriptProperties for ONBOARDING_PACKAGE tokens that are still ACTIVE
 * but whose matter already has both AGREEMENT and POA signed.
 * Returns array of { token, matterId, entry }.
 */
function crm_findStaleOnboardingPackageTokens_() {
  var all = PropertiesService.getScriptProperties().getProperties();
  var prefix = CRM_SIGN_TOKEN_PREFIX;
  var stale = [];
  for (var key in all) {
    if (!key.startsWith(prefix)) continue;
    if (key.endsWith("_expires")) continue;
    var entry;
    try { entry = JSON.parse(all[key]); } catch (e) { continue; }
    if (!entry || entry.kind !== "ONBOARDING_PACKAGE") continue;
    if (entry.status !== CRM_SIGN_TOKEN_STATUS_ACTIVE) continue;
    if (!entry.matterId) continue;
    // Expired tokens are not "stale active" — already harmless
    if (crm_isTokenExpired(entry)) continue;
    // Stale = still active, but signing is already fully done
    if (crm_hasBothSignedOnboardingDocs_(entry.matterId)) {
      stale.push({ token: entry.token, matterId: entry.matterId, entry: entry });
    }
  }
  return stale;
}

// ─── Core audit function ───────────────────────────────────────────────────

/**
 * Scans CRM data for known consistency problems. Performs NO writes.
 *
 * opts.skipActivities  — set true to skip the Activities scan (faster but misses NMR dupes)
 *
 * Returns:
 * {
 *   ok: true,
 *   auditOnly: true,
 *   summary: { leadsChecked, mattersChecked, docsChecked, activitiesChecked, issuesFound, repairableCount },
 *   findings: [ { type, severity, repairable, repaired:false, ...context } ]
 * }
 */
function crm_auditConsistency_(opts) {
  opts = opts || {};
  var findings = [];

  // Load raw sheet data (Activities capped at 1000 most-recent rows for performance)
  var allLeads      = crm_auditLoadRows_("LEADS");
  var allMatters    = crm_auditLoadRows_("MATTERS");
  var allDocs       = crm_auditLoadRows_("DOCUMENTS");
  var allActs       = opts.skipActivities ? [] : crm_auditLoadRows_("ACTIVITIES", 1000);

  // Index matters by CLIENT_ID for O(1) lookup per lead
  var mattersByClient = {};
  for (var i = 0; i < allMatters.length; i++) {
    var m = allMatters[i];
    var cid = String(m.CLIENT_ID || "").trim();
    if (!cid) continue;
    if (!mattersByClient[cid]) mattersByClient[cid] = [];
    mattersByClient[cid].push(m);
  }

  // ── Checks 1 & 2: Lead-level + provisional linkage ──────────────────────
  for (var i = 0; i < allLeads.length; i++) {
    var lead      = allLeads[i];
    var leadId    = String(lead.LEAD_ID || "").trim();
    var clientId  = String(lead.CLIENT_ID || "").trim();
    var status    = String(lead.STATUS || "").toUpperCase().trim();
    if (!leadId) continue;

    var isConverted    = status === "CONVERTED";
    var hasRealClient  = clientId && !crm_isProvisionalClientId_(clientId);
    var hasProvisional = crm_isProvisionalClientId_(clientId);

    // 1a. STATUS=CONVERTED but no real CLIENT_ID
    if (isConverted && !hasRealClient) {
      findings.push({
        type:      "LEAD_CONVERTED_BAD_CLIENT",
        severity:  "HIGH",
        leadId:    leadId,
        clientId:  clientId || null,
        detail:    "Lead STATUS=CONVERTED but CLIENT_ID is " + (clientId ? "provisional (" + clientId + ")" : "empty"),
        repairable: true,
        repaired:   false,
      });
    }

    // 1b. Has real CLIENT_ID but STATUS is still an onboarding transient state
    // (only repair onboarding-derived statuses; "NEW" with a client is ambiguous legacy data)
    if (hasRealClient && (status === "ONBOARDING" || status === "PENDING_SIGNATURE")) {
      findings.push({
        type:          "LEAD_STATUS_NOT_CONVERTED",
        severity:      "MEDIUM",
        leadId:        leadId,
        clientId:      clientId,
        currentStatus: status,
        detail:        "Lead has real CLIENT_ID=" + clientId + " but STATUS=" + status,
        repairable:    true,
        repaired:      false,
      });
    }

    // 2. Lead has real CLIENT_ID but matters or docs still carry provisional leadId
    if (hasRealClient) {
      var provisionalMatters = mattersByClient[leadId] || [];
      for (var j = 0; j < provisionalMatters.length; j++) {
        var pm = provisionalMatters[j];
        var pmId = String(pm.MATTER_ID || "").trim();
        if (!pmId) continue;

        findings.push({
          type:         "MATTER_PROVISIONAL_AFTER_CONVERSION",
          severity:     "HIGH",
          leadId:       leadId,
          realClientId: clientId,
          matterId:     pmId,
          detail:       "Matter " + pmId + " CLIENT_ID still provisional: " + leadId,
          repairable:   true,
          repaired:     false,
        });

        // Check docs for this matter still pointing to leadId
        var provisionalDocs = allDocs.filter(function (d) {
          return String(d.MATTER_ID || "") === pmId && String(d.CLIENT_ID || "") === leadId;
        });
        if (provisionalDocs.length > 0) {
          findings.push({
            type:         "DOCS_PROVISIONAL_AFTER_CONVERSION",
            severity:     "MEDIUM",
            leadId:       leadId,
            realClientId: clientId,
            matterId:     pmId,
            docIds:       provisionalDocs.map(function (d) { return String(d.DOC_ID || ""); }),
            count:        provisionalDocs.length,
            detail:       provisionalDocs.length + " doc(s) still have provisional CLIENT_ID for matter " + pmId,
            repairable:   true,
            repaired:     false,
          });
        }
      }
    }
  }

  // ── Check 3: Multiple ONBOARDING matters per client/lead context ─────────
  var onboardingByContext = {};
  for (var i = 0; i < allMatters.length; i++) {
    var m = allMatters[i];
    if (String(m.STAGE || "").toUpperCase() !== "ONBOARDING") continue;
    var key = String(m.CLIENT_ID || "").trim();
    if (!key) continue;
    if (!onboardingByContext[key]) onboardingByContext[key] = [];
    onboardingByContext[key].push(m);
  }
  for (var ctxKey in onboardingByContext) {
    if (onboardingByContext[ctxKey].length > 1) {
      findings.push({
        type:      "MULTIPLE_ONBOARDING_MATTERS",
        severity:  "MEDIUM",
        clientId:  ctxKey,
        matterIds: onboardingByContext[ctxKey].map(function (m) { return String(m.MATTER_ID || ""); }),
        count:     onboardingByContext[ctxKey].length,
        detail:    onboardingByContext[ctxKey].length + " ONBOARDING matters for client/lead context: " + ctxKey,
        repairable: false,
        repaired:   false,
      });
    }
  }

  // ── Check 4: Duplicate signed outputs per matter/type ────────────────────
  var signedByKey = {};
  for (var i = 0; i < allDocs.length; i++) {
    var d = allDocs[i];
    var dtype = String(d.TYPE || "").toUpperCase().trim();
    if (dtype !== "SIGNED_AGREEMENT" && dtype !== "SIGNED_POA") continue;
    if (String(d.STATUS || "").toUpperCase() !== "SIGNED") continue;
    var matterId = String(d.MATTER_ID || "").trim();
    if (!matterId) continue;
    var sKey = matterId + "||" + dtype;
    if (!signedByKey[sKey]) signedByKey[sKey] = [];
    signedByKey[sKey].push(d);
  }
  for (var sKey in signedByKey) {
    if (signedByKey[sKey].length > 1) {
      var skParts = sKey.split("||");
      findings.push({
        type:       "DUPLICATE_SIGNED_OUTPUT",
        severity:   "LOW",
        matterId:   skParts[0],
        signedType: skParts[1],
        docIds:     signedByKey[sKey].map(function (d) { return String(d.DOC_ID || ""); }),
        count:      signedByKey[sKey].length,
        detail:     signedByKey[sKey].length + " " + skParts[1] + " rows for matter " + skParts[0],
        repairable: false,
        repaired:   false,
      });
    }
  }

  // ── Check 5: Duplicate NEXT_MODULE_READY per matter/module ───────────────
  var nxByKey = {};
  for (var i = 0; i < allActs.length; i++) {
    var a = allActs[i];
    if (String(a.ACTION || "") !== "NEXT_MODULE_READY") continue;
    var meta = {};
    try { meta = typeof a.META_JSON === "string" ? JSON.parse(a.META_JSON) : (a.META_JSON || {}); } catch (e) {}
    var module = String(meta.module || "").trim();
    if (!module || !a.MATTER_ID) continue;
    var nKey = String(a.MATTER_ID).trim() + "||" + module;
    nxByKey[nKey] = (nxByKey[nKey] || 0) + 1;
  }
  for (var nKey in nxByKey) {
    if (nxByKey[nKey] > 1) {
      var nkParts = nKey.split("||");
      findings.push({
        type:      "DUPLICATE_NEXT_MODULE_READY",
        severity:  "LOW",
        matterId:  nkParts[0],
        module:    nkParts[1],
        count:     nxByKey[nKey],
        detail:    "NEXT_MODULE_READY fired " + nxByKey[nKey] + "× for matter " + nkParts[0] + " module " + nkParts[1],
        repairable: false,
        repaired:   false,
      });
    }
  }

  // ── Check 6: Stale active onboarding package tokens ──────────────────────
  var staleTokens = crm_findStaleOnboardingPackageTokens_();
  for (var i = 0; i < staleTokens.length; i++) {
    findings.push({
      type:      "STALE_ACTIVE_ONBOARDING_TOKEN",
      severity:  "LOW",
      matterId:  staleTokens[i].matterId,
      token:     staleTokens[i].token,
      detail:    "ONBOARDING_PACKAGE token still ACTIVE for fully-signed matter: " + staleTokens[i].matterId,
      repairable: true,
      repaired:   false,
    });
  }

  return {
    ok:        true,
    auditOnly: true,
    summary: {
      leadsChecked:      allLeads.length,
      mattersChecked:    allMatters.length,
      docsChecked:       allDocs.length,
      activitiesChecked: allActs.length,
      issuesFound:       findings.length,
      repairableCount:   findings.filter(function (f) { return f.repairable; }).length,
    },
    findings: findings,
  };
}

// ─── Safe repair function ──────────────────────────────────────────────────

/**
 * Runs the audit then applies deterministic safe repairs.
 * Repair is idempotent: running it twice converges to the same state.
 *
 * Returns:
 * {
 *   ok: true,
 *   auditOnly: false,
 *   summary: { ... },
 *   findings: [ ... ],   // all findings including non-repairable
 *   repaired: [ ... ],   // findings that were successfully repaired
 *   skipped:  [ ... ],   // findings not repaired (non-repairable or errored)
 * }
 */
function crm_repairConsistency_(opts) {
  opts = opts || {};
  var audit    = crm_auditConsistency_(opts);
  var repaired = [];
  var skipped  = [];

  for (var i = 0; i < audit.findings.length; i++) {
    var f = audit.findings[i];

    if (!f.repairable) {
      skipped.push(f);
      continue;
    }

    try {

          // ── LEAD_STATUS_NOT_CONVERTED ──────────────────────────────────────────
      if (f.type === "LEAD_STATUS_NOT_CONVERTED") {
        crm_updateLeadStatus(f.leadId, "CONVERTED");
        crm_logActivity({
          action:   "CONSISTENCY_REPAIR",
          message:  "Repaired lead STATUS → CONVERTED for lead " + f.leadId +
                    " (already has real CLIENT_ID: " + f.clientId + ")",
          clientId: f.clientId,
          meta:     { repairType: f.type, leadId: f.leadId, clientId: f.clientId },
        });
        f.repaired = true;
        repaired.push(f);
        continue;
      }

      // ── LEAD_CONVERTED_BAD_CLIENT (SAFE RECOVERY) ─────────────────────────
      if (f.type === "LEAD_CONVERTED_BAD_CLIENT") {
        var leadBad = crm_getLeadById(f.leadId);
        if (leadBad) {
          var currentClientId = String(leadBad.CLIENT_ID || "").trim();
          var realClientId = "";

          // If a real client is already linked somehow, reuse it.
          if (currentClientId && !crm_isProvisionalClientId_(currentClientId)) {
            realClientId = currentClientId;
          } else {
            // Create a real client directly from lead data.
            var addRes = crm_addClient({
              fullName: String(leadBad.FULL_NAME || "").trim(),
              phone: String(leadBad.PHONE || "").trim(),
              email: String(leadBad.EMAIL || "").trim(),
              idType: String(leadBad.ID_TYPE || "").trim(),
              idNumber: String(leadBad.ID_NUMBER || "").trim(),
              address: String(leadBad.ADDRESS || "").trim(),
              source: "CONSISTENCY_REPAIR",
              status: "NEW",
              owner: String(leadBad.ASSIGNED_TO || getActiveUserEmail_() || "").trim(),
            });
            realClientId = addRes.clientId || "";
          }

          if (!realClientId) {
            throw new Error("Failed to recover real client for lead " + f.leadId);
          }

          // Update lead row: CLIENT_ID + STATUS=CONVERTED
          var cLead = cfg_();
          var shLead = crm_getSpreadsheet_().getSheetByName(cLead.SHEETS.LEADS);
          if (shLead) {
            var headersLead = shLead.getRange(1, 1, 1, shLead.getLastColumn()).getValues()[0];
            var iLeadId = headersLead.indexOf("LEAD_ID");
            var iClient = headersLead.indexOf("CLIENT_ID");
            var iStatus = headersLead.indexOf("STATUS");
            var iUpdated = headersLead.indexOf("UPDATED_AT");

            if (iLeadId >= 0) {
              var leadValues = shLead.getRange(2, iLeadId + 1, shLead.getLastRow() - 1, 1).getValues();
              for (var lr = 0; lr < leadValues.length; lr++) {
                if (String(leadValues[lr][0] || "").trim() === String(f.leadId)) {
                  var row = lr + 2;
                  if (iClient >= 0) shLead.getRange(row, iClient + 1).setValue(realClientId);
                  if (iStatus >= 0) shLead.getRange(row, iStatus + 1).setValue("CONVERTED");
                  if (iUpdated >= 0) shLead.getRange(row, iUpdated + 1).setValue(nowIso_());
                  break;
                }
              }
            }
          }

          // Reattach any provisional matters/docs anchored to the leadId
          var badMatters = crm_findMattersByClientId(f.leadId) || [];
          for (var bm = 0; bm < badMatters.length; bm++) {
            var badMatterId = String(badMatters[bm].MATTER_ID || "").trim();
            if (!badMatterId) continue;
            crm_setMatterField(badMatterId, "CLIENT_ID", realClientId);
            crm_reattachMatterDocsToClient_(badMatterId, f.leadId, realClientId);
          }

          crm_logActivity({
            action:   "CONSISTENCY_REPAIR",
            message:  "Recovered missing client for converted lead " + f.leadId,
            clientId: realClientId,
            meta:     {
              repairType: f.type,
              leadId: f.leadId,
              clientId: realClientId
            },
          });

          f.repaired = true;
          repaired.push(f);
          continue;
        }
      }

      // ── MATTER_PROVISIONAL_AFTER_CONVERSION ───────────────────────────────
      if (f.type === "MATTER_PROVISIONAL_AFTER_CONVERSION") {
        crm_setMatterField(f.matterId, "CLIENT_ID", f.realClientId);
        crm_logActivity({
          action:   "CONSISTENCY_REPAIR",
          message:  "Repaired provisional matter CLIENT_ID: " + f.matterId + " → " + f.realClientId,
          clientId: f.realClientId,
          matterId: f.matterId,
          meta:     { repairType: f.type, leadId: f.leadId, matterId: f.matterId, realClientId: f.realClientId },
        });
        f.repaired = true;
        repaired.push(f);
        continue;
      }

      // ── DOCS_PROVISIONAL_AFTER_CONVERSION ─────────────────────────────────
      if (f.type === "DOCS_PROVISIONAL_AFTER_CONVERSION") {
        crm_reattachMatterDocsToClient_(f.matterId, f.leadId, f.realClientId);
        crm_logActivity({
          action:   "CONSISTENCY_REPAIR",
          message:  "Reattached " + f.count + " provisional doc(s) for matter " + f.matterId + " → " + f.realClientId,
          clientId: f.realClientId,
          matterId: f.matterId,
          meta:     { repairType: f.type, leadId: f.leadId, matterId: f.matterId, realClientId: f.realClientId, count: f.count },
        });
        f.repaired = true;
        repaired.push(f);
        continue;
      }

      // ── STALE_ACTIVE_ONBOARDING_TOKEN ─────────────────────────────────────
      if (f.type === "STALE_ACTIVE_ONBOARDING_TOKEN") {
        crm_expireSignToken(f.token);
        crm_logActivity({
          action:   "TOKEN_EXPIRED_REPAIR",
          message:  "Expired stale ONBOARDING_PACKAGE token for fully-signed matter " + f.matterId,
          matterId: f.matterId,
          meta:     { repairType: f.type, token: f.token, matterId: f.matterId },
        });
        f.repaired = true;
        repaired.push(f);
        continue;
      }

      // Unknown repairable type — should not happen; skip safely
      skipped.push(f);

    } catch (e) {
      f.repairError = e.message;
      skipped.push(f);
      try {
        crm_logActivity({
          action:   "CONSISTENCY_WARNING",
          message:  "Repair failed for " + f.type + ": " + e.message,
          matterId: f.matterId || "",
          meta:     { repairType: f.type, error: e.message, leadId: f.leadId || "", matterId: f.matterId || "" },
        });
      } catch (logErr) { /* ignore log failure — don't obscure repair error */ }
    }
  }

  return {
    ok:        true,
    auditOnly: false,
    summary:   audit.summary,
    findings:  audit.findings,
    repaired:  repaired,
    skipped:   skipped,
  };
}

// ─── Menu-callable wrappers ────────────────────────────────────────────────

function crm_runConsistencyAudit() {
  var result  = crm_auditConsistency_({});
  var s       = result.summary;
  Logger.log(JSON.stringify(result, null, 2));

  var msg = "Audit complete.\n\n"
    + "Scanned: " + s.leadsChecked + " leads · " + s.mattersChecked + " matters · "
    + s.docsChecked + " docs · " + s.activitiesChecked + " activities.\n\n"
    + "Issues found: " + s.issuesFound
    + "  (" + s.repairableCount + " auto-repairable).\n\n"
    + "See View › Logs for full details.";

  try {
    SpreadsheetApp.getUi().alert("CRM Consistency Audit", msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) { /* called outside spreadsheet context — results already in Logger */ }

  return result;
}

function crm_runConsistencyRepair() {
  var result = crm_repairConsistency_({});
  Logger.log(JSON.stringify(result, null, 2));

  var msg = "Repair complete.\n\n"
    + "Issues found: " + result.summary.issuesFound + "\n"
    + "Repaired:     " + result.repaired.length + "\n"
    + "Skipped (report-only or error): " + result.skipped.length + "\n\n"
    + "See View › Logs for full details.";

  try {
    SpreadsheetApp.getUi().alert("CRM Consistency Repair", msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) { /* called outside spreadsheet context */ }

  return result;
}