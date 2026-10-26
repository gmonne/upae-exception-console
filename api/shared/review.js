// Logic for the project-manager review page: which findings a PM may see, and how a response is validated.
const LISTS = { findings: "WeeklyFindings", jobs: "JobAssignments", history: "FindingHistory" };
const RESPONSES = ["Correct as reported", "Needs correction", "Question"];
const STATUS_DONE = "Resolved", STATUS_SENT = "Sent to PM";

const clean = (v) => String(v == null ? "" : v).replace(/\s+/g, " ").trim();
// 23-430~0328 and 25-615-0103 belong to jobs 23-430 and 25-615
function jobBase(j) { j = clean(j); const m = j.match(/^(\d{2}-\d{3})(?:[~-].*)?$/); return m ? m[1] : j; }
function jobBasesOf(text) {
  const out = [], seen = {};
  String(text || "").split("; ").forEach((x) => {
    const j = x.replace(/ \(.*$/, "").trim();
    if (!j || j === "(no job)") return;
    const b = jobBase(j); if (!seen[b]) { seen[b] = 1; out.push(b); }
  });
  return out;
}
function pmJobs(assignmentItems, email) {
  const e = clean(email).toLowerCase(), out = {};
  (assignmentItems || []).forEach((it) => {
    const f = it.fields || {};
    if (f.Active === false) return;
    if (clean(f.PMEmail).toLowerCase() !== e) return;
    const b = jobBase(f.Title); if (b) out[b] = { job: b, jobName: clean(f.JobName), pmName: clean(f.PMName) };
  });
  return out;
}
function shape(it, mine) {
  const f = it.fields || {};
  const bases = jobBasesOf(f.Jobs);
  return {
    id: it.id, key: f.FindingKey || "", weekStart: f.WeekStart || "", weekNo: f.WeekNo != null ? Number(f.WeekNo) : null,
    employeeId: f.EmployeeID || "", employeeName: f.EmployeeName || "", union: f.UnionName || "", day: f.WorkDate || "", kind: f.Kind || "",
    ruleName: f.RuleName || "", broken: f.RuleBroken || "", reported: f.Reported || "", expected: f.Expected || "",
    jobs: f.Jobs || "", myJobs: bases.filter((b) => mine[b]), status: f.Status || "Open", sentDate: f.SentDate || "", payrollNote: f.Note || "",
    pmResponse: f.PMResponse || "", pmNote: f.PMNote || "", pmBy: f.PMRespondedBy || "", pmDate: f.PMRespondedDate || "",
    ai: f.AIExplanation ? { explanation: f.AIExplanation, suggestedAction: f.AISuggestedAction || "", confidence: f.AIConfidence || "" } : null,
  };
}
function canSee(it, mine) {
  const f = it.fields || {};
  if (!jobBasesOf(f.Jobs).some((b) => mine[b])) return false;
  const st = f.Status || "Open";
  return st === STATUS_SENT;       // only records Payroll has sent to PMs; resolved or not-yet-sent ones are not shown
}
function selectFor(findingItems, mine) {
  return (findingItems || []).filter((it) => canSee(it, mine)).map((it) => shape(it, mine))
    .sort((a, b) => (b.weekStart || "").localeCompare(a.weekStart || "") || (a.myJobs[0] || "").localeCompare(b.myJobs[0] || "") || a.employeeId.localeCompare(b.employeeId) || (a.day || "").localeCompare(b.day || ""));
}
function validateResponse(r) {
  const response = clean(r && r.response), note = String(r && r.note || "").trim().slice(0, 1000);
  if (!RESPONSES.includes(response)) throw Object.assign(new Error("Unknown response."), { status: 400 });
  if (response !== "Correct as reported" && !note) throw Object.assign(new Error("Please add a note for '" + response + "'."), { status: 400 });
  return { response, note };
}
module.exports = { LISTS, RESPONSES, STATUS_DONE, STATUS_SENT, jobBase, jobBasesOf, pmJobs, selectFor, shape, canSee, validateResponse };
