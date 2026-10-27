// Prompt, input trimming, output validation and the Claude call for the Weekly Analyst Summary (Phase 2).
// All figures are computed in the browser from the same data as the dashboard. The model only writes the narrative,
// and every number in its reply is checked against the figures it was given.
const PROMPT_VERSION = "summary-v2";

const SYSTEM = [
  "You are the weekly analyst assistant inside UPAE, a payroll accuracy tool for a construction company.",
  "A deterministic rules engine has already checked each union employee's daily hours (RT regular, OT 1.5x, DT 2x) against the union rule. You never change its results and never calculate pay.",
  "You write a short briefing for a payroll administrator about ONE week, using only the figures and findings supplied in the user message.",
  "",
  "Rules you must follow:",
  "1. Use ONLY the supplied data. Every number you write must appear exactly in the supplied data. Do not add, subtract, average or compute percentages yourself; use the percentages supplied.",
  "2. Compare with the prior weeks only when prior weeks are supplied; say so plainly if there are none.",
  "3. A day with OT or DT hours and no RT hours (flagged noRtDay) can be a legitimate second shift or work on another job. Describe such items as 'to verify', never as errors. If source is 'summary', the report has NO job numbers, so say in dataLimits that job detail is not available. If source is 'detail', each candidate carries the jobs worked that day and the owning project manager, and 'byJob' and 'pmWaiting' are supplied: use them.",
  "3b. Jobs and project managers: mention a job or PM only if it appears in byJob, pmWaiting or a candidate. In jobFocus, list up to 5 jobs that need attention first (most open findings, or findings waiting on a PM), each with one short sentence. Say plainly when jobs have no PM assigned (jobsWithoutPm) because nobody can be asked yet. A PM answer of 'Needs correction' or 'Question' is waiting on payroll; 'Correct as reported' still needs payroll to resolve it.",
  "4. Do not mention pay rates or dollar amounts. Do not give legal advice. Do not guess causes you cannot see in the data.",
  "5. In riskRanked, use only finding keys from the supplied candidates list, most important first, and explain each in one short sentence. Prefer items that are still Open, repeat employees, and large differences.",
  "6. Keep everything short, concrete and written for a busy reader.",
  "",
  "Reply with a single JSON object and nothing else, with exactly these keys:",
  '{"headline": string (1-2 sentences), "keyChanges": string[] (2-5 short items: what stands out, and how the week compares with prior weeks if available), "riskRanked": [{"key": string, "risk": "high"|"medium"|"low", "reason": string}] (up to 8), "suggestedFollowUps": string[] (up to 5 concrete next steps for payroll or project managers), "jobFocus": [{"job": string (a job number from byJob), "note": string}] (up to 5, empty array if source is summary), "dataLimits": string (1-2 sentences on what this data cannot show), "confidence": "low"|"medium"|"high"}',
].join("\n");

const s = (v, n) => String(v == null ? "" : v).slice(0, n || 200);
const num = (v) => { const x = Number(v); return isFinite(x) ? Math.round(x * 100) / 100 : 0; };

// Keep only the expected fields, with size limits. Throws on bad input.
function trimInput(body) {
  if (!body || typeof body !== "object" || !body.week || !body.totals) { const e = new Error("A week with totals is required."); e.status = 400; throw e; }
  const list = (a, n) => (Array.isArray(a) ? a.slice(0, n) : []);
  return {
    week: { weekNo: body.week.weekNo == null ? null : num(body.week.weekNo), start: s(body.week.start, 10), end: s(body.week.end, 10) },
    totals: { employeesInFile: num(body.totals.employeesInFile), rt: num(body.totals.rt), ot: num(body.totals.ot), dt: num(body.totals.dt), totalHours: num(body.totals.totalHours), otSharePct: num(body.totals.otSharePct), dtSharePct: num(body.totals.dtSharePct) },
    coverage: body.coverage ? { employeesChecked: num(body.coverage.employeesChecked), daysChecked: num(body.coverage.daysChecked), daysMatching: num(body.coverage.daysMatching), checkedHoursPct: num(body.coverage.checkedHoursPct), unionsSkipped: list(body.coverage.unionsSkipped, 12).map((u) => ({ union: s(u.union, 80), employees: num(u.employees) })) } : null,
    findings: {
      total: num(body.findings && body.findings.total),
      byKind: list(body.findings && body.findings.byKind, 10).map((k) => ({ kind: s(k.kind, 40), count: num(k.count) })),
      byStatus: { open: num(body.findings && body.findings.byStatus && body.findings.byStatus.open), sentToPm: num(body.findings && body.findings.byStatus && body.findings.byStatus.sentToPm), resolved: num(body.findings && body.findings.byStatus && body.findings.byStatus.resolved) },
      ruleDifferencesByUnion: list(body.findings && body.findings.ruleDifferencesByUnion, 8).map((u) => ({ union: s(u.union, 80), differences: num(u.differences), daysChecked: num(u.daysChecked) })),
      employeesWithMostDifferences: list(body.findings && body.findings.employeesWithMostDifferences, 8).map((e) => ({ employee: s(e.employee, 100), differences: num(e.differences) })),
      noRtDayDifferences: num(body.findings && body.findings.noRtDayDifferences),
      resolvedReasons: list(body.findings && body.findings.resolvedReasons, 8).map((r) => ({ reason: s(r.reason, 60), count: num(r.count) })),
    },
    source: body.source === "detail" ? "detail" : "summary",
    byJob: list(body.byJob, 12).map((j) => ({ job: s(j.job, 20), name: s(j.name, 80), pm: s(j.pm, 80), findings: num(j.findings), open: num(j.open), sentToPm: num(j.sentToPm), pmAnswered: num(j.pmAnswered), resolved: num(j.resolved) })),
    pmWaiting: list(body.pmWaiting, 10).map((p) => ({ pm: s(p.pm, 80), waiting: num(p.waiting), answered: num(p.answered) })),
    jobsWithoutPm: num(body.jobsWithoutPm),
    pmAnswers: { correct: num(body.pmAnswers && body.pmAnswers.correct), needsCorrection: num(body.pmAnswers && body.pmAnswers.needsCorrection), question: num(body.pmAnswers && body.pmAnswers.question) },
    priorWeeks: list(body.priorWeeks, 6).map((w) => ({ weekNo: w.weekNo == null ? null : num(w.weekNo), start: s(w.start, 10), totalHours: num(w.totalHours), ruleDifferences: num(w.ruleDifferences), otSharePct: w.otSharePct == null ? null : num(w.otSharePct) })),
    candidates: list(body.candidates, 25).map((c) => ({ key: s(c.key, 80), employee: s(c.employee, 100), union: s(c.union, 80), day: s(c.day, 20), reported: s(c.reported, 60), ruleSays: s(c.ruleSays, 60), kind: s(c.kind, 30), status: s(c.status, 20), jobs: s(c.jobs, 120), pm: s(c.pm, 80), pmResponse: s(c.pmResponse, 30), noRtDay: !!c.noRtDay, hasAiNote: !!c.hasAiNote, repeatEmployeeCount: num(c.repeatEmployeeCount) })),
  };
}

function buildUserMessage(input) {
  return "Write the weekly briefing. Data (JSON):\n" + JSON.stringify(input, null, 2);
}

// Every number mentioned in the narrative must be a number that was supplied (dates count as supplied numbers).
function numbersIn(text) { return (String(text).match(/\d+(?:,\d{3})*(?:\.\d+)?/g) || []).map((x) => String(Number(x.replace(/,/g, "")))); }
function allowedNumbers(input) { return new Set(numbersIn(JSON.stringify(input))); }
function checkNumbers(out, input) {
  const ok = allowedNumbers(input), bad = [];
  const texts = [out.headline, out.dataLimits].concat(out.keyChanges, out.suggestedFollowUps, out.riskRanked.map((r) => r.reason), (out.jobFocus || []).map((j) => j.note));
  texts.forEach((t) => numbersIn(t).forEach((n) => { if (!ok.has(n)) bad.push(n); }));
  return bad;
}

function parseOutput(text, input) {
  const m = String(text || "").match(/\{[\s\S]*\}/);
  if (!m) throw new Error("No JSON in reply");
  const o = JSON.parse(m[0]);
  const conf = ["low", "medium", "high"].includes(o.confidence) ? o.confidence : null;
  if (typeof o.headline !== "string" || !o.headline.trim()) throw new Error("headline missing");
  if (!conf) throw new Error("confidence invalid");
  const strs = (a, n, len) => (Array.isArray(a) ? a.filter((x) => typeof x === "string" && x.trim()).slice(0, n).map((x) => x.slice(0, len)) : []);
  const keys = new Set((input && input.candidates || []).map((c) => c.key));
  const risk = (Array.isArray(o.riskRanked) ? o.riskRanked : []).filter((r) => r && keys.has(r.key) && typeof r.reason === "string").slice(0, 8)
    .map((r) => ({ key: r.key, risk: ["high", "medium", "low"].includes(r.risk) ? r.risk : "medium", reason: r.reason.slice(0, 300) }));
  const jobSet = new Set((input && input.byJob || []).map((j) => j.job));
  const jobFocus = (Array.isArray(o.jobFocus) ? o.jobFocus : []).filter((j) => j && jobSet.has(String(j.job)) && typeof j.note === "string").slice(0, 5).map((j) => ({ job: String(j.job), note: j.note.slice(0, 300) }));
  const out = {
    jobFocus, headline: o.headline.slice(0, 500), keyChanges: strs(o.keyChanges, 5, 400), riskRanked: risk, suggestedFollowUps: strs(o.suggestedFollowUps, 5, 400),
    dataLimits: typeof o.dataLimits === "string" ? o.dataLimits.slice(0, 500) : "", confidence: conf,
  };
  if (input) { const bad = checkNumbers(out, input); if (bad.length) throw new Error("figures not in the data: " + bad.slice(0, 5).join(", ")); }
  return out;
}

async function callClaude(input, opts) {
  async function once(extra) {
    const res = await (opts.fetch || fetch)("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: { "content-type": "application/json", "x-api-key": opts.apiKey, "anthropic-version": "2023-06-01" },
      body: JSON.stringify({ model: opts.model, max_tokens: 4096, system: SYSTEM, messages: [{ role: "user", content: buildUserMessage(input) + (extra || "") }] }),
    });
    const body = await res.json().catch(() => ({}));
    if (!res.ok) { const e = new Error("The AI service returned " + res.status + (body.error && body.error.message ? ": " + body.error.message : "")); e.status = res.status === 429 ? 429 : 502; throw e; }
    const txt = (body.content || []).map((c) => c.text || "").join("");
    if (body.stop_reason === "max_tokens") throw new Error("reply was cut off (too long)");
    return txt;
  }
  let text;
  try { text = await once(""); return parseOutput(text, input); }
  catch (e1) {
    if (e1.status === 429 || (e1.status && e1.status >= 500)) throw e1;
    text = await once("\n\nKeep every field short (the whole reply under 350 words). "+"\n\nYour previous reply was rejected (" + String(e1.message).slice(0, 120) + "). Reply with ONLY the JSON object, using only numbers that appear in the data.");
    try { return parseOutput(text, input); }
    catch (e2) { const e = new Error("The summary could not be checked against the data (" + String(e2.message).slice(0, 100) + "). Please try again."); e.status = 502; throw e; }
  }
}

// Used when no API key is configured yet, so the screen can be tested end to end.
function stubOutput(input) {
  const f = input.findings;
  return {
    headline: "[Sample summary: the AI key is not set up yet.] Week " + (input.week.weekNo || "") + " has " + f.total + " findings.",
    keyChanges: ["Sample text. Once the API key is added, this describes what stands out this week and how it compares with prior weeks."],
    riskRanked: input.candidates.slice(0, 3).map((c) => ({ key: c.key, risk: "medium", reason: "Sample ranking for " + c.employee + "." })),
    suggestedFollowUps: ["Review the open findings and email the project managers."],
    jobFocus: [], dataLimits: "The weekly report has no job numbers or clock times.", confidence: "low",
  };
}

module.exports = { PROMPT_VERSION, SYSTEM, trimInput, buildUserMessage, parseOutput, checkNumbers, callClaude, stubOutput };
