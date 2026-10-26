/* ============================================================================
   UPAE — SharePoint / Microsoft Graph Integration Layer
   ----------------------------------------------------------------------------
   Replaces the mock-up's hardcoded rules[]/unions[]/laborTypes[] arrays with
   live reads/writes against the 5 SharePoint lists via Microsoft Graph.

   REQUIRES: MSAL Browser library loaded before this file, e.g. in <head>:
     <script src="https://alcdn.msauth.net/browser/2.38.3/js/msal-browser.min.js"></script>

   ASSUMPTIONS TO VERIFY IN AZURE PORTAL (flagged per Gus's question, Sept 23):
     1. App registration "Certified Payroll App" is platform type SPA (Single-page
        application), NOT "Web". If it's "Web", auth below will fail silently or
        redirect incorrectly — fix in Azure Portal > App registrations > this app
        > Authentication > Platform configurations.
     2. Redirect URI registered under that SPA platform matches window.location.origin
        of wherever this mock-up is actually hosted (the Azure Static Web Apps URL).
     3. API permission Sites.ReadWrite.All (delegated) has admin consent granted.
        If using Sites.Selected instead, the app must additionally be granted
        access to this specific site by a SharePoint/Global admin — a permission
        scope alone is not enough for Sites.Selected.
     4. Signed-in users need at least Contribute permission on the SharePoint site
        itself for writes to succeed — Graph will return 403 otherwise, and the app
        can't grant this from the browser side.
   ============================================================================ */

const UPAE_CONFIG = {
  tenantId: "c62585ee-57f1-43a7-80b5-8efc4d57e464",
  clientId: "b9e533a5-4da9-4999-80aa-7fc5bf2cd50a", // "UPAE Rule Library" app registration, Sept 23 2026
  siteHostname: "cacindinc.sharepoint.com",
  sitePath: "/sites/CACContractManagement",
  // Internal list names as created in SharePoint (must match exactly — Graph
  // resolves lists by display name OR internal GUID; display name is used here
  // for readability, but if list renames ever happen this breaks, so consider
  // swapping to list GUIDs once you have them from the Graph calls below).
  lists: {
    unions: "Unions",
    laborTypes: "LaborTypes",
    rules: "Rules",
    ruleLaborTypes: "RuleLaborTypes",
    ruleExamples: "RuleExamples",
    employees: "Employees",
    weeklyHours: "WeeklyHours",
    weeklyFindings: "WeeklyFindings",
    projectManagers: "ProjectManagers",
    jobAssignments: "JobAssignments",
    aiLog: "AILog",
    findingHistory: "FindingHistory",
    weeklySummaries: "WeeklySummaries",
  },
};

const msalConfig = {
  auth: {
    clientId: UPAE_CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${UPAE_CONFIG.tenantId}`,
    redirectUri: window.location.origin, // must match the SPA redirect URI registered in Azure AD
  },
  cache: {
    cacheLocation: "sessionStorage", // safer than localStorage for a shared/public machine
  },
};

const graphScopes = ["Sites.ReadWrite.All"]; // switch to your Sites.Selected scope string if used instead

const msalInstance = new msal.PublicClientApplication(msalConfig);
let upaeSiteId = null; // resolved once per session, cached below

/* ---------------------------- AUTH ---------------------------- */

async function upaeSignIn() {
  try {
    const result = await msalInstance.loginPopup({ scopes: graphScopes });
    msalInstance.setActiveAccount(result.account);
    return result.account;
  } catch (err) {
    console.error("UPAE sign-in failed:", err);
    throw err;
  }
}

async function upaeGetToken() {
  const account = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
  if (!account) {
    const signedIn = await upaeSignIn();
    return upaeGetToken(); // retry now that an account exists
  }
  try {
    const result = await msalInstance.acquireTokenSilent({ scopes: graphScopes, account });
    return result.accessToken;
  } catch (err) {
    // silent acquisition can fail if consent hasn't been granted yet, or token expired mid-session
    const result = await msalInstance.acquireTokenPopup({ scopes: graphScopes });
    return result.accessToken;
  }
}

async function upaeGraphFetch(path, options = {}) {
  const token = await upaeGetToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`Graph API ${options.method || "GET"} ${path} failed: ${res.status} ${body}`);
  }
  return res.status === 204 ? null : res.json();
}

/* ------------------------ SITE / LIST RESOLUTION ------------------------ */

async function upaeGetSiteId() {
  if (upaeSiteId) return upaeSiteId;
  const site = await upaeGraphFetch(`/sites/${UPAE_CONFIG.siteHostname}:${UPAE_CONFIG.sitePath}`);
  upaeSiteId = site.id;
  return upaeSiteId;
}

async function upaeListItemsUrl(listKey, extra = "") {
  const siteId = await upaeGetSiteId();
  const listName = UPAE_CONFIG.lists[listKey];
  return `/sites/${siteId}/lists/${listName}/items${extra}`;
}

/* Resolves a list item's SharePoint numeric id by its Title. Needed anywhere
   we write a Lookup column: Microsoft Graph will NOT accept a plain string
   value for a Lookup field's own internal name (e.g. { Rule: "Local 14..." })
   — it silently drops it instead of erroring, leaving the lookup blank. The
   only way to set a Lookup column via Graph is <InternalName>LookupId: <id>,
   where <id> is the target item's SharePoint id, not its display text. This
   resolves that id by fetching the target list and matching on Title.
   Client-side match (not $filter) to sidestep OData quoting/indexing quirks
   on small lists like these. */
async function upaeGetListItemId(listKey, title) {
  // Nested $expand=fields(select=Title) is rejected by this tenant's Graph
  // endpoint for SharePoint list items ("Term '(select=Title)' is not valid"),
  // so expand the full fields object instead of trying to select just Title.
  const url = await upaeListItemsUrl(listKey, "?$expand=fields&$top=1000");
  const data = await upaeGraphFetch(url);
  const match = data.value.find((item) => item.fields.Title === title);
  return match ? parseInt(match.id, 10) : null;
}

/* ------------------------------ READS ------------------------------
   Each returns a plain-JS-array shaped to match the mock-up's existing
   in-memory data structures, so render functions need minimal changes. */

async function upaeFetchUnions() {
  const url = await upaeListItemsUrl("unions", "?expand=fields&$top=200");
  const data = await upaeGraphFetch(url);
  return data.value.map((item) => ({
    id: item.id,
    name: item.fields.Title,
    unionCode: item.fields.UnionCode || "",
    cbaReference: item.fields.CBAReference || "",
  }));
}

async function upaeFetchLaborTypes(liveUnions) {
  // Union is a Lookup column — confirmed (same bug as Rule/LaborType elsewhere)
  // that Graph does NOT return it as a plain string on fields.Union; it comes
  // back as fields.UnionLookupId (a number). Map that id back to the union's
  // name using the already-fetched Unions list instead of trusting fields.Union.
  const url = await upaeListItemsUrl(
    "laborTypes",
    "?expand=fields&$top=500"
  );
  const data = await upaeGraphFetch(url);
  const unionIdToName = {};
  (liveUnions || []).forEach((u) => { unionIdToName[String(u.id)] = u.name; });
  return data.value.map((item) => ({
    id: item.id,
    code: item.fields.LaborCode,
    name: item.fields.Title,
    unionName: unionIdToName[String(item.fields.UnionLookupId)] || null,
    sourceId: item.fields.SourceID || "",
  }));
}

async function upaeFetchRules() {
  const url = await upaeListItemsUrl("rules", "?expand=fields&$top=200");
  const data = await upaeGraphFetch(url);
  return data.value.map((item) => {
    const f = item.fields;
    return {
      id: f.RuleID,
      spItemId: item.id, // keep the raw SharePoint item id for updates
      name: f.Title,
      status: (f.Status || "").toLowerCase().includes("structured")
        ? "structured"
        : (f.Status || "").toLowerCase().includes("ambiguous")
        ? "ambiguous"
        : (f.Status || "").toLowerCase().includes("auto-resolved")
        ? "ambiguous"
        : "draft",
      statusLabel: f.Status || "Draft - Prose Only",
      cba: f.CBAReference || "—",
      verified: f.VerifiedBy && f.VerifiedDate ? `${f.VerifiedBy.LookupValue || f.VerifiedBy}, ${f.VerifiedDate}` : "Not yet verified",
      codes: f.CodesNote || "",
      fields: {
        "Work week start day": f.WorkWeekStartDay || "",
        "Standard workday (hrs)": f.StandardWorkdayHrs != null ? String(f.StandardWorkdayHrs) : "",
        "Regular threshold": f.RegularThreshold || "",
        "Regular applicable days": f.RegularApplicableDays || "",
        "OT 1.5 trigger type": f.OT15TriggerType || "",
        "OT 1.5 threshold": f.OT15Threshold || "",
        "OT 1.5 applicable days": f.OT15ApplicableDays || "",
        "OT 1.5 rate mult.": f.OT15RateMult != null ? String(f.OT15RateMult) : "",
        "OT 1.5 benefit mult.": f.OT15BenefitMult != null ? String(f.OT15BenefitMult) : "",
        "OT 2x trigger type": f.OT2TriggerType || "",
        "OT 2x threshold": f.OT2Threshold || "",
        "OT 2x applicable days": f.OT2ApplicableDays || "",
        "OT 2x rate mult.": f.OT2RateMult != null ? String(f.OT2RateMult) : "",
        "OT 2x benefit mult.": f.OT2BenefitMult != null ? String(f.OT2BenefitMult) : "",
        "Night trigger type": f.NightTriggerType || "",
        "Night trigger condition": f.NightTriggerCondition || "",
        "Night window start": f.NightWindowStart || "",
        "Night window end": f.NightWindowEnd || "",
        "Night applicability": f.NightApplicability || "",
        "Night tier 1 threshold": f.NightTier1Threshold || "",
        "Night tier 1 rate mult.": f.NightTier1RateMult != null ? String(f.NightTier1RateMult) : "",
        "Night tier 1 benefit mult.": f.NightTier1BenefitMult != null ? String(f.NightTier1BenefitMult) : "",
        "Night tier 2 threshold": f.NightTier2Threshold || "",
        "Night tier 2 rate mult.": f.NightTier2RateMult != null ? String(f.NightTier2RateMult) : "",
        "Night tier 2 benefit mult.": f.NightTier2BenefitMult != null ? String(f.NightTier2BenefitMult) : "",
        "Precedence rule": f.PrecedenceRule || "",
      },
      examples: {}, // populated separately via upaeFetchRuleExamples(ruleId) — not eagerly loaded for all rules to limit initial payload
    };
  });
}

async function upaeFetchRuleLaborTypes(ruleTitle, liveLaborTypes) {
  // Rule/LaborType here are SharePoint Lookup columns. Graph never returns a
  // Lookup as a flat string on its own internal name — it comes back as
  // "<Field>LookupId" (a number) with "<Field>" itself either absent or a
  // nested {LookupId, LookupValue} object, never a plain string. So we can't
  // filter/read by comparing item.fields.Rule === ruleTitle (always false).
  // Instead resolve the rule's numeric id and match on RuleLookupId, then map
  // the LaborType's numeric SharePoint id back to its short code (r.laborTypeIds
  // is expected to hold codes like "OP-1", matching laborTypes[].code).
  const ruleId = await upaeGetListItemId("rules", ruleTitle);
  if (!ruleId) return [];
  const url = await upaeListItemsUrl(
    "ruleLaborTypes",
    `?expand=fields&$top=500`
  );
  const data = await upaeGraphFetch(url);
  const laborTypeIdToCode = {};
  (liveLaborTypes || []).forEach((lt) => { laborTypeIdToCode[String(lt.id)] = lt.code; });
  return data.value
    .filter((item) => String(item.fields.RuleLookupId) === String(ruleId))
    .map((item) => laborTypeIdToCode[String(item.fields.LaborTypeLookupId)])
    .filter((code) => code != null);
}

async function upaeFetchRuleExamples(ruleTitle) {
  // Same Lookup-read issue as above: match on RuleLookupId, not fields.Rule.
  const ruleId = await upaeGetListItemId("rules", ruleTitle);
  if (!ruleId) return {};
  const url = await upaeListItemsUrl("ruleExamples", `?expand=fields&$top=500`);
  const data = await upaeGraphFetch(url);
  const rows = data.value.filter((item) => String(item.fields.RuleLookupId) === String(ruleId));
  const grouped = {};
  rows.forEach((item) => {
    const cat = item.fields.Category;
    if (!grouped[cat]) grouped[cat] = [];
    grouped[cat].push({ ok: !!item.fields.IsCorrect, text: item.fields.ExampleText, spItemId: item.id });
  });
  return grouped;
}

/* ------------------------------ WRITES ------------------------------ */

async function upaeCreateUnion(name, unionCode = "", cbaReference = "") {
  const url = await upaeListItemsUrl("unions");
  return upaeGraphFetch(url, {
    method: "POST",
    body: JSON.stringify({ fields: { Title: name, UnionCode: unionCode, CBAReference: cbaReference } }),
  });
}

async function upaeDeleteUnion(spItemId) {
  const siteId = await upaeGetSiteId();
  await upaeGraphFetch(`/sites/${siteId}/lists/${UPAE_CONFIG.lists.unions}/items/${spItemId}`, {
    method: "DELETE",
  });
}

async function upaeUpdateUnion(spItemId, name, unionCode, cbaReference) {
  const siteId = await upaeGetSiteId();
  await upaeGraphFetch(`/sites/${siteId}/lists/${UPAE_CONFIG.lists.unions}/items/${spItemId}/fields`, {
    method: "PATCH",
    body: JSON.stringify({ Title: name, UnionCode: unionCode, CBAReference: cbaReference }),
  });
}

async function upaeCreateLaborType(name, code, unionName = null, sourceId = "") {
  const fields = { Title: name, LaborCode: code, SourceID: sourceId };
  if (unionName) {
    const unionId = await upaeGetListItemId("unions", unionName);
    if (unionId) fields.UnionLookupId = unionId;
  }
  const url = await upaeListItemsUrl("laborTypes");
  return upaeGraphFetch(url, { method: "POST", body: JSON.stringify({ fields }) });
}

async function upaeUpdateLaborTypeUnion(spItemId, unionName) {
  const siteId = await upaeGetSiteId();
  const unionId = unionName ? await upaeGetListItemId("unions", unionName) : null;
  await upaeGraphFetch(
    `/sites/${siteId}/lists/${UPAE_CONFIG.lists.laborTypes}/items/${spItemId}/fields`,
    { method: "PATCH", body: JSON.stringify({ UnionLookupId: unionId }) }
  );
}

async function upaeDeleteLaborType(spItemId) {
  const siteId = await upaeGetSiteId();
  await upaeGraphFetch(`/sites/${siteId}/lists/${UPAE_CONFIG.lists.laborTypes}/items/${spItemId}`, {
    method: "DELETE",
  });
}

// fieldsPayload keys should be the SharePoint internal column names (e.g. RuleID,
// Status, OT2RateMult, etc.) — see sharepoint-list-schema.md for the exact list.
async function upaeCreateRule(fieldsPayload) {
  const url = await upaeListItemsUrl("rules");
  return upaeGraphFetch(url, { method: "POST", body: JSON.stringify({ fields: fieldsPayload }) });
}

async function upaeUpdateRule(spItemId, fieldsPayload) {
  const siteId = await upaeGetSiteId();
  await upaeGraphFetch(`/sites/${siteId}/lists/${UPAE_CONFIG.lists.rules}/items/${spItemId}/fields`, {
    method: "PATCH",
    body: JSON.stringify(fieldsPayload),
  });
}

async function upaeCreateRuleLaborType(ruleTitle, laborTypeTitle) {
  const [ruleId, laborTypeId] = await Promise.all([
    upaeGetListItemId("rules", ruleTitle),
    upaeGetListItemId("laborTypes", laborTypeTitle),
  ]);
  if (!ruleId || !laborTypeId) {
    throw new Error(
      `Could not find a matching Rules/LaborTypes item for "${ruleTitle}" / "${laborTypeTitle}" — check the Title spelling matches exactly.`
    );
  }
  const url = await upaeListItemsUrl("ruleLaborTypes");
  return upaeGraphFetch(url, {
    method: "POST",
    body: JSON.stringify({
      fields: { Title: `${ruleTitle} - ${laborTypeTitle}`, RuleLookupId: ruleId, LaborTypeLookupId: laborTypeId },
    }),
  });
}

async function upaeCreateRuleExample(ruleTitle, category, isCorrect, exampleText) {
  const ruleId = await upaeGetListItemId("rules", ruleTitle);
  if (!ruleId) {
    throw new Error(`Could not find a Rules item titled "${ruleTitle}" — the example was not saved.`);
  }
  const url = await upaeListItemsUrl("ruleExamples");
  return upaeGraphFetch(url, {
    method: "POST",
    body: JSON.stringify({
      fields: {
        Title: `${ruleTitle} - ${category} - ${isCorrect ? "Correct" : "Incorrect"}`,
        RuleLookupId: ruleId,
        Category: category,
        IsCorrect: isCorrect,
        ExampleText: exampleText,
      },
    }),
  });
}

async function upaeUpdateRuleExample(spItemId, category, isCorrect, exampleText) {
  const siteId = await upaeGetSiteId();
  await upaeGraphFetch(
    `/sites/${siteId}/lists/${UPAE_CONFIG.lists.ruleExamples}/items/${spItemId}/fields`,
    {
      method: "PATCH",
      body: JSON.stringify({
        Category: category,
        IsCorrect: isCorrect,
        ExampleText: exampleText,
      }),
    }
  );
}

async function upaeDeleteRuleExample(spItemId) {
  const siteId = await upaeGetSiteId();
  await upaeGraphFetch(`/sites/${siteId}/lists/${UPAE_CONFIG.lists.ruleExamples}/items/${spItemId}`, {
    method: "DELETE",
  });
}

/* ------------------------------ EMPLOYEES ------------------------------
   The Employees list holds the employee master (from the B2W Employee Listing
   spreadsheet). Plain-text columns only - no Lookup columns - so none of the
   Lookup read/write quirks above apply. Internal column names:
     Title (= "Last, First"), EmployeeID, LastName, FirstName, BusinessUnit,
     DefaultLaborCode, DefaultLaborName, Role
   An employee's union is NOT stored: it is derived in the app from
   DefaultLaborCode -> LaborTypes.Union, so renaming/re-mapping a union
   never leaves stale employee rows behind. */

function upaeEmployeeFromItem(item) {
  const f = item.fields || {};
  return {
    spItemId: item.id,
    id: f.EmployeeID || "",
    last: f.LastName || "",
    first: f.FirstName || "",
    businessUnit: f.BusinessUnit || "",
    laborCode: f.DefaultLaborCode || "",
    laborName: f.DefaultLaborName || "",
    role: f.Role || "",
    unionOverride: f.UnionOverride || "",
    pmEmail: f.PMEmail || "",
  };
}

function upaeEmployeeToFields(e) {
  return {
    Title: [e.last, e.first].filter(Boolean).join(", ") || e.id,
    EmployeeID: e.id,
    LastName: e.last || "",
    FirstName: e.first || "",
    BusinessUnit: e.businessUnit || "",
    DefaultLaborCode: e.laborCode || "",
    DefaultLaborName: e.laborName || "",
    Role: e.role || "",
    // only sent when set (or being cleared), so saving works before the UnionOverride column exists
    ...(e.unionOverride || e.clearOverride ? { UnionOverride: e.unionOverride || "" } : {}),
    ...(e.pmEmail || e.clearPm ? { PMEmail: e.pmEmail || "" } : {}),
  };
}

async function upaeFetchEmployees() {
  // Follows @odata.nextLink so lists larger than one page are read in full.
  let url = await upaeListItemsUrl("employees", "?expand=fields&$top=500");
  const out = [];
  let guard = 0;
  while (url && guard++ < 100) {
    const data = await upaeGraphFetch(url);
    (data.value || []).forEach((item) => out.push(upaeEmployeeFromItem(item)));
    const next = data["@odata.nextLink"];
    // nextLink is absolute; upaeGraphFetch wants a path relative to /v1.0
    url = next ? next.replace("https://graph.microsoft.com/v1.0", "") : null;
  }
  return out;
}

async function upaeCreateEmployee(e) {
  const url = await upaeListItemsUrl("employees");
  const item = await upaeGraphFetch(url, {
    method: "POST",
    body: JSON.stringify({ fields: upaeEmployeeToFields(e) }),
  });
  return item; // item.id is the new SharePoint item id
}

async function upaeUpdateEmployee(spItemId, e) {
  const siteId = await upaeGetSiteId();
  await upaeGraphFetch(`/sites/${siteId}/lists/${UPAE_CONFIG.lists.employees}/items/${spItemId}/fields`, {
    method: "PATCH",
    body: JSON.stringify(upaeEmployeeToFields(e)),
  });
}

// Remember (or clear) an employee's usual project manager without touching any other column
async function upaeSetEmployeePM(spItemId, email) {
  const siteId = await upaeGetSiteId();
  await upaeGraphFetch(`/sites/${siteId}/lists/${UPAE_CONFIG.lists.employees}/items/${spItemId}/fields`, { method: "PATCH", body: JSON.stringify({ PMEmail: email || "" }) });
}
async function upaeDeleteEmployee(spItemId) {
  const siteId = await upaeGetSiteId();
  await upaeGraphFetch(`/sites/${siteId}/lists/${UPAE_CONFIG.lists.employees}/items/${spItemId}`, {
    method: "DELETE",
  });
}

/* Bulk create/update through Graph's $batch endpoint (max 20 requests per
   batch), so a first import of ~1,100 employees is ~57 calls instead of 1,100.
   ops: [{ kind: "create"|"update", employee, spItemId? }]
   Returns one result per op, in order: { ok, spItemId?, error? }.
   Throttled (429/503/504) sub-requests are retried up to 4 times, honouring
   Retry-After. onProgress(doneCount, total) is called after each batch. */
async function upaeBatchList(listKey, ops, toFields, onProgress) {
  const siteId = await upaeGetSiteId();
  const base = `/sites/${siteId}/lists/${UPAE_CONFIG.lists[listKey]}/items`;
  const results = new Array(ops.length);
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  let done = 0;
  for (let start = 0; start < ops.length; start += 20) {
    let pending = [];
    for (let i = start; i < Math.min(start + 20, ops.length); i++) pending.push(i);
    for (let attempt = 0; attempt < 5 && pending.length; attempt++) {
      const requests = pending.map((i) => {
        const op = ops[i];
        return op.kind === "create"
          ? { id: String(i), method: "POST", url: base, headers: { "Content-Type": "application/json" }, body: { fields: toFields(op) } }
          : { id: String(i), method: "PATCH", url: `${base}/${op.spItemId}/fields`, headers: { "Content-Type": "application/json" }, body: toFields(op) };
      });
      const res = await upaeGraphFetch("/$batch", { method: "POST", body: JSON.stringify({ requests }) });
      const retry = [];
      let waitMs = 0;
      (res.responses || []).forEach((r) => {
        const i = parseInt(r.id, 10);
        if (r.status >= 200 && r.status < 300) {
          results[i] = { ok: true, spItemId: r.body && r.body.id ? r.body.id : ops[i].spItemId };
        } else if ([429, 503, 504].includes(r.status) && attempt < 4) {
          retry.push(i);
          const ra = r.headers && (r.headers["Retry-After"] || r.headers["retry-after"]);
          waitMs = Math.max(waitMs, (parseInt(ra, 10) || 2) * 1000);
        } else {
          const msg = r.body && r.body.error && r.body.error.message ? r.body.error.message : "HTTP " + r.status;
          results[i] = { ok: false, error: msg };
        }
      });
      // Any request missing from the response is treated as failed, not silently dropped.
      pending.forEach((i) => {
        if (!results[i] && !retry.includes(i)) results[i] = { ok: false, error: "No response from SharePoint for this row" };
      });
      pending = retry;
      if (pending.length) await sleep(waitMs);
    }
    pending.forEach((i) => { if (!results[i]) results[i] = { ok: false, error: "SharePoint kept throttling this row" }; });
    done += Math.min(20, ops.length - start);
    if (typeof onProgress === "function") onProgress(done, ops.length);
  }
  return results;
}

async function upaeBatchEmployees(ops, onProgress) {
  return upaeBatchList("employees", ops, (op) => upaeEmployeeToFields(op.employee), onProgress);
}

// ---- Weekly hours: one row per employee per worked day ----
function upaeWeeklyHoursFromItem(item) {
  const f = item.fields || {};
  const id = f.EmployeeID || "";
  return {
    spItemId: item.id,
    key: String(id).trim().toUpperCase() + "|" + (f.WorkDate || ""),
    employeeId: id,
    employeeName: f.EmployeeName || "",
    workDate: f.WorkDate || "",
    weekStart: f.WeekStart || "",
    weekNo: f.WeekNo != null ? Number(f.WeekNo) : null,
    rt: Number(f.RT) || 0,
    ot: Number(f.OT) || 0,
    dt: Number(f.DT) || 0,
  };
}
function upaeWeeklyHoursToFields(r) {
  return {
    Title: r.employeeId + " " + r.workDate,
    EmployeeID: r.employeeId,
    EmployeeName: r.employeeName || "",
    WorkDate: r.workDate,
    WeekStart: r.weekStart,
    WeekNo: r.weekNo,
    RT: r.rt,
    OT: r.ot,
    DT: r.dt,
  };
}
async function upaeFetchWeeklyHours() {
  let url = await upaeListItemsUrl("weeklyHours", "?expand=fields&$top=500");
  const out = [];
  let guard = 0;
  while (url && guard++ < 400) {
    const data = await upaeGraphFetch(url);
    (data.value || []).forEach((item) => out.push(upaeWeeklyHoursFromItem(item)));
    const next = data["@odata.nextLink"];
    url = next ? next.replace("https://graph.microsoft.com/v1.0", "") : null;
  }
  return out;
}
async function upaeBatchWeeklyHours(ops, onProgress) {
  return upaeBatchList("weeklyHours", ops, (op) => upaeWeeklyHoursToFields(op.row), onProgress);
}

/* ------------------------------ LOAD + WIRE-UP ------------------------------
   Called by the sign-in button (see index.html/upae-mockup.html). Fetches all
   5 lists, then hands the results to window.upaeApplyLiveData — a bridge
   function defined inside upae-mockup.html's own <script> IIFE — because
   rules/unions/laborTypes live in that closure, not on window, so this file
   can't reassign them directly. See the wiring guide, section 3. */

async function upaeLoadAllData() {
  const statusEl = document.getElementById('upaeSignInStatus');
  function setStatus(msg) { if (statusEl) statusEl.textContent = msg; }

  setStatus('Loading unions...');
  const liveUnions = await upaeFetchUnions();

  setStatus('Loading labor types...');
  const liveLaborTypes = await upaeFetchLaborTypes(liveUnions);

  setStatus('Loading rules...');
  const liveRules = await upaeFetchRules();

  setStatus('Loading rule details (labor codes + examples)...');
  for (const r of liveRules) {
    r.laborTypeIds = await upaeFetchRuleLaborTypes(r.name, liveLaborTypes);
    r.examples = await upaeFetchRuleExamples(r.name);
  }

  // Employees are loaded last and NEVER block sign-in: if the Employees list has
  // not been created in SharePoint yet, the rest of the app still works and the
  // Employees screen explains what is missing.
  let liveEmployees = null;
  let employeesError = null;
  setStatus('Loading employees...');
  try {
    liveEmployees = await upaeFetchEmployees();
  } catch (err) {
    console.warn('UPAE: could not load the Employees list', err);
    employeesError = err.message || String(err);
  }

  if (typeof window.upaeApplyLiveData !== 'function') {
    throw new Error(
      'window.upaeApplyLiveData is not defined — make sure sharepoint-integration.js ' +
      'loads BEFORE the mock-up\'s own <script> block that defines it, and that this ' +
      'function runs after that block has executed (e.g. triggered by a button click, ' +
      'not at page-load time before the IIFE below it has run).'
    );
  }
  window.upaeApplyLiveData(liveUnions, liveLaborTypes, liveRules);
  if (typeof window.upaeApplyEmployees === 'function') {
    window.upaeApplyEmployees(liveEmployees, employeesError);
  }
  setStatus('Loaded.');
  if (typeof window.upaeDashRefresh === 'function') window.upaeDashRefresh();
}

// ---- Weekly findings (what the rule check flagged, plus review status) ----
function upaeFindingFromItem(item) {
  const f = item.fields || {};
  return {
    spItemId: item.id,
    key: f.FindingKey || "",
    weekStart: f.WeekStart || "",
    weekNo: f.WeekNo != null ? Number(f.WeekNo) : null,
    id: f.EmployeeID || "",
    name: f.EmployeeName || "",
    union: f.UnionName || "",
    iso: f.WorkDate || "",
    kind: f.Kind || "",
    ruleId: f.RuleID || "",
    ruleName: f.RuleName || "",
    broken: f.RuleBroken || "",
    reported: f.Reported || "",
    expected: f.Expected || "",
    status: f.Status || "Open",
    sentTo: f.SentTo || "",
    sentDate: f.SentDate || "",
    note: f.Note || "",
    aiExplanation: f.AIExplanation || "",
    aiLikelyCause: f.AILikelyCause || "",
    aiSuggestedAction: f.AISuggestedAction || "",
    aiEvidence: f.AIEvidence || "",
    aiConfidence: f.AIConfidence || "",
    aiModel: f.AIModel || "",
    aiDate: f.AIDate || "",
    aiFeedback: f.AIFeedback || "",
    jobs: f.Jobs || "",
    resReason: f.ResolutionReason || "",
    resolvedBy: f.ResolvedBy || "",
    resolvedDate: f.ResolvedDate || "",
    modified: item.lastModifiedDateTime || "",
    modifiedBy: (item.lastModifiedBy && item.lastModifiedBy.user && (item.lastModifiedBy.user.email || item.lastModifiedBy.user.displayName)) || "",
  };
}
function upaeFindingToFields(r) {
  return {
    Title: r.id + " " + (r.iso || "") + " " + r.kind,
    FindingKey: r.key,
    WeekStart: r.weekStart,
    WeekNo: r.weekNo,
    EmployeeID: r.id,
    EmployeeName: r.name || "",
    UnionName: r.union || "",
    WorkDate: r.iso || "",
    Kind: r.kind,
    RuleID: r.ruleId || "",
    RuleName: r.ruleName || "",
    RuleBroken: r.broken || "",
    Reported: r.reported || "",
    Expected: r.expected || "",
    Status: r.status || "Open",
    SentTo: r.sentTo || "",
    SentDate: r.sentDate || "",
    Note: r.note || "",
    // AI columns are only sent once an explanation exists, so saving still works before those columns are added
    ...(r.aiExplanation ? {
      AIExplanation: r.aiExplanation,
      AILikelyCause: r.aiLikelyCause || "",
      AISuggestedAction: r.aiSuggestedAction || "",
      AIEvidence: r.aiEvidence || "",
      AIConfidence: r.aiConfidence || "",
      AIModel: r.aiModel || "",
      AIDate: r.aiDate || "",
      AIFeedback: r.aiFeedback || "",
    } : {}),
    // the Jobs column is only written once a job is known
    ...(r.jobs ? { Jobs: r.jobs } : {}),
    // resolution fields are only sent once a reason exists
    ...(r.resReason ? { ResolutionReason: r.resReason, ResolvedBy: r.resolvedBy || "", ResolvedDate: r.resolvedDate || "" } : {}),
  };
}
async function upaeFetchFindings() {
  let url = await upaeListItemsUrl("weeklyFindings", "?expand=fields&$top=500");
  const out = [];
  let guard = 0;
  while (url && guard++ < 400) {
    const data = await upaeGraphFetch(url);
    (data.value || []).forEach((item) => out.push(upaeFindingFromItem(item)));
    const next = data["@odata.nextLink"];
    url = next ? next.replace("https://graph.microsoft.com/v1.0", "") : null;
  }
  return out;
}
// Who last changed these rows, and when (used to warn before overwriting somebody else's change).
async function upaeFetchFindingStamps(ids) {
  const siteId = await upaeGetSiteId();
  const base = `/sites/${siteId}/lists/${UPAE_CONFIG.lists.weeklyFindings}/items`;
  const out = {};
  for (let start = 0; start < ids.length; start += 20) {
    const chunk = ids.slice(start, start + 20);
    const requests = chunk.map((id, i) => ({ id: String(i), method: "GET", url: `${base}/${id}?$select=id,lastModifiedDateTime,lastModifiedBy` }));
    const res = await upaeGraphFetch("/$batch", { method: "POST", body: JSON.stringify({ requests }) });
    (res.responses || []).forEach((r) => {
      const id = chunk[parseInt(r.id, 10)];
      if (r.status >= 200 && r.status < 300 && r.body) {
        const u = r.body.lastModifiedBy && r.body.lastModifiedBy.user;
        out[id] = { modified: r.body.lastModifiedDateTime || "", modifiedBy: u ? (u.email || u.displayName || "") : "" };
      } else if (r.status === 404) out[id] = { gone: true };
    });
  }
  return out;
}

// ---- Change history: one row each time a finding's status changes ----
async function upaeLogHistory(entries) {
  const ops = entries.map((e) => ({ kind: "create", row: e }));
  return upaeBatchList("findingHistory", ops, (op) => ({
    Title: (op.row.key || "").slice(0, 250),
    FindingKey: op.row.key || "",
    WeekStart: op.row.weekStart || "",
    EmployeeID: op.row.id || "",
    EmployeeName: op.row.name || "",
    FromStatus: op.row.from || "",
    ToStatus: op.row.to || "",
    Detail: op.row.detail || "",
    ChangedBy: op.row.by || "",
  }));
}
async function upaeFetchHistory() {
  let url = await upaeListItemsUrl("findingHistory", "?expand=fields&$top=500");
  const out = [];
  let guard = 0;
  while (url && guard++ < 400) {
    const data = await upaeGraphFetch(url);
    (data.value || []).forEach((item) => {
      const f = item.fields || {};
      out.push({ key: f.FindingKey || "", weekStart: f.WeekStart || "", id: f.EmployeeID || "", name: f.EmployeeName || "", from: f.FromStatus || "", to: f.ToStatus || "", detail: f.Detail || "", by: f.ChangedBy || "", at: item.createdDateTime || "" });
    });
    const next = data["@odata.nextLink"];
    url = next ? next.replace("https://graph.microsoft.com/v1.0", "") : null;
  }
  return out;
}
async function upaeBatchFindings(ops, onProgress) {
  return upaeBatchList("weeklyFindings", ops, (op) => upaeFindingToFields(op.row), onProgress);
}

// ---- Project managers (name + email) ----
async function upaeFetchProjectManagers() {
  const url = await upaeListItemsUrl("projectManagers", "?expand=fields&$top=500");
  const data = await upaeGraphFetch(url);
  return (data.value || []).map((item) => ({ spItemId: item.id, name: (item.fields || {}).Title || "", email: (item.fields || {}).Email || "" }));
}
async function upaeCreateProjectManager(pm) {
  const url = await upaeListItemsUrl("projectManagers");
  return upaeGraphFetch(url, { method: "POST", body: JSON.stringify({ fields: { Title: pm.name, Email: pm.email } }) });
}
async function upaeUpdateProjectManager(spItemId, pm) {
  const siteId = await upaeGetSiteId();
  await upaeGraphFetch(`/sites/${siteId}/lists/${UPAE_CONFIG.lists.projectManagers}/items/${spItemId}/fields`, { method: "PATCH", body: JSON.stringify({ Title: pm.name, Email: pm.email }) });
}
async function upaeDeleteProjectManager(spItemId) {
  const siteId = await upaeGetSiteId();
  await upaeGraphFetch(`/sites/${siteId}/lists/${UPAE_CONFIG.lists.projectManagers}/items/${spItemId}`, { method: "DELETE" });
}

// ---- Job assignments: which project manager owns which job (one row per job) ----
async function upaeFetchJobAssignments() {
  let url = await upaeListItemsUrl("jobAssignments", "?expand=fields&$top=500");
  const out = [];
  while (url) {
    const data = await upaeGraphFetch(url);
    (data.value || []).forEach((item) => {
      const f = item.fields || {};
      out.push({ spItemId: item.id, job: f.Title || "", jobName: f.JobName || "", pmEmail: f.PMEmail || "", pmName: f.PMName || "", active: f.Active !== false });
    });
    url = data["@odata.nextLink"] || null;
  }
  return out;
}
async function upaeCreateJobAssignment(a) {
  const url = await upaeListItemsUrl("jobAssignments");
  return upaeGraphFetch(url, { method: "POST", body: JSON.stringify({ fields: { Title: a.job, JobName: a.jobName || "", PMEmail: a.pmEmail || "", PMName: a.pmName || "", Active: true } }) });
}
async function upaeUpdateJobAssignment(spItemId, a) {
  const siteId = await upaeGetSiteId();
  await upaeGraphFetch(`/sites/${siteId}/lists/${UPAE_CONFIG.lists.jobAssignments}/items/${spItemId}/fields`, { method: "PATCH", body: JSON.stringify({ JobName: a.jobName || "", PMEmail: a.pmEmail || "", PMName: a.pmName || "" }) });
}

// ---- Send an email as the signed-in user (needs the Mail.Send delegated permission) ----
async function upaeGetMailToken() {
  const account = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
  if (!account) { await upaeSignIn(); return upaeGetMailToken(); }
  const scopes = ["Mail.Send"];
  try {
    return (await msalInstance.acquireTokenSilent({ scopes, account })).accessToken;
  } catch (err) {
    return (await msalInstance.acquireTokenPopup({ scopes })).accessToken;
  }
}
async function upaeSendMail(toList, ccList, subject, htmlBody) {
  const token = await upaeGetMailToken();
  const rec = (a) => ({ emailAddress: { address: a } });
  const res = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
    method: "POST",
    headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({
      message: { subject, body: { contentType: "HTML", content: htmlBody }, toRecipients: toList.map(rec), ccRecipients: (ccList || []).map(rec) },
      saveToSentItems: true,
    }),
  });
  if (!res.ok) {
    const t = await res.text().catch(() => "");
    throw new Error("Email was not sent: " + res.status + " " + t.slice(0, 300));
  }
}

// ---- AI environment: calls the Azure Function in this Static Web App (/api/ai/*) ----
async function upaeGetIdToken() {
  const account = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
  if (!account) { await upaeSignIn(); return upaeGetIdToken(); }
  let result;
  try { result = await msalInstance.acquireTokenSilent({ scopes: graphScopes, account }); }
  catch (err) { result = await msalInstance.acquireTokenPopup({ scopes: graphScopes }); }
  return result.idToken;
}
async function upaeAiPost(path, payload) {
  const idToken = await upaeGetIdToken();
  let lastErr = null;
  // One quiet retry: a bare 5xx from Azure (cold start, worker restart) usually clears on the second try.
  for (let attempt = 0; attempt < 2; attempt++) {
    let res;
    try {
      res = await fetch(path, {
        method: "POST",
        // Static Web Apps rewrites the standard Authorization header on /api calls, so the sign-in token travels in its own header.
        headers: { "X-UPAE-Token": "Bearer " + idToken, "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
    } catch (err) {
      lastErr = new Error("Could not reach the AI service. It may not be deployed yet.");
      if (attempt === 0) { await new Promise((r) => setTimeout(r, 1500)); continue; }
      throw lastErr;
    }
    const text = await res.text();
    let body = null;
    try { body = JSON.parse(text); } catch (e) { /* not JSON */ }
    if (res.ok && body) return body;
    const serverMsg = body && body.error;
    lastErr = new Error(serverMsg || (res.status === 404 ? "The AI service is not deployed yet (404)." : res.status >= 500 ? "The AI service had a temporary problem (" + res.status + "). Please try again." : "The AI service returned " + res.status + "."));
    // retry only bare server errors (no message from our own code)
    if (attempt === 0 && res.status >= 500 && !serverMsg) { await new Promise((r) => setTimeout(r, 1500)); continue; }
    throw lastErr;
  }
  throw lastErr || new Error("The AI request failed.");
}
async function upaeAiExplain(payload) { return upaeAiPost("/api/ai/explain", payload); }
async function upaeAiSummary(payload) { return upaeAiPost("/api/ai/summary", payload); }

// ---- Weekly summaries (Phase 2): one row each time a summary is saved ----
async function upaeSaveSummary(r) {
  const url = await upaeListItemsUrl("weeklySummaries");
  return upaeGraphFetch(url, { method: "POST", body: JSON.stringify({ fields: {
    Title: "Week " + (r.weekNo || "") + " " + (r.weekStart || ""),
    WeekStart: r.weekStart || "", WeekNo: r.weekNo == null ? null : r.weekNo,
    Headline: (r.headline || "").slice(0, 500), SummaryJSON: r.summaryJson || "",
    Model: r.model || "", PromptVersion: r.promptVersion || "", Confidence: r.confidence || "",
    GeneratedBy: r.generatedBy || "", Feedback: "",
  } }) });
}
async function upaeFetchSummaries() {
  let url = await upaeListItemsUrl("weeklySummaries", "?expand=fields&$top=500");
  const out = [];
  let guard = 0;
  while (url && guard++ < 100) {
    const data = await upaeGraphFetch(url);
    (data.value || []).forEach((item) => {
      const f = item.fields || {};
      out.push({ spItemId: item.id, weekStart: f.WeekStart || "", weekNo: f.WeekNo != null ? Number(f.WeekNo) : null, headline: f.Headline || "", json: f.SummaryJSON || "", model: f.Model || "",
        promptVersion: f.PromptVersion || "", confidence: f.Confidence || "", generatedBy: f.GeneratedBy || "", feedback: f.Feedback || "", at: item.createdDateTime || "" });
    });
    const next = data["@odata.nextLink"];
    url = next ? next.replace("https://graph.microsoft.com/v1.0", "") : null;
  }
  return out;
}
async function upaeUpdateSummaryFeedback(spItemId, feedback) {
  const siteId = await upaeGetSiteId();
  await upaeGraphFetch(`/sites/${siteId}/lists/${UPAE_CONFIG.lists.weeklySummaries}/items/${spItemId}/fields`, { method: "PATCH", body: JSON.stringify({ Feedback: feedback }) });
}
// AILog: audit trail of every AI answer, plus the reviewer's feedback (the learning signal for later phases)
async function upaeCreateAILog(r) {
  const url = await upaeListItemsUrl("aiLog");
  return upaeGraphFetch(url, {
    method: "POST",
    body: JSON.stringify({ fields: {
      Title: (r.purpose || "explain") + " " + (r.findingKey || ""),
      Purpose: r.purpose || "explain",
      FindingKey: r.findingKey || "",
      WeekStart: r.weekStart || "",
      UserEmail: r.userEmail || "",
      Model: r.model || "",
      PromptVersion: r.promptVersion || "",
      Confidence: r.confidence || "",
      Output: r.output || "",
      Feedback: "",
    } }),
  });
}
async function upaeUpdateAILogFeedback(spItemId, feedback) {
  const siteId = await upaeGetSiteId();
  await upaeGraphFetch(`/sites/${siteId}/lists/${UPAE_CONFIG.lists.aiLog}/items/${spItemId}/fields`, { method: "PATCH", body: JSON.stringify({ Feedback: feedback }) });
}
function upaeCurrentUserEmail() {
  const a = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
  return a ? (a.username || "") : "";
}

/* Wires the #upaeSignInBtn button (added in upae-mockup.html's <body>) to the
   full sign-in + load flow, and hides the sign-in gate overlay on success. */
document.addEventListener('DOMContentLoaded', function () {
  const btn = document.getElementById('upaeSignInBtn');
  const gate = document.getElementById('upaeSignInGate');
  const statusEl = document.getElementById('upaeSignInStatus');
  if (!btn) return; // this page doesn't have the sign-in gate — nothing to wire
  btn.addEventListener('click', async function () {
    btn.disabled = true;
    btn.textContent = 'Signing in...';
    try {
      await upaeSignIn();
      await upaeLoadAllData();
      if (gate) gate.style.display = 'none';
    } catch (err) {
      console.error('UPAE sign-in/load failed:', err);
      if (statusEl) statusEl.textContent = 'Failed: ' + err.message;
      btn.disabled = false;
      btn.textContent = 'Retry sign-in';
    }
  });
});
