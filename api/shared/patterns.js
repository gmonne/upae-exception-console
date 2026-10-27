// Phase 3 pattern detection: the browser finds statistical outliers in code; this module only asks the model to interpret those
// outliers. The model cannot add outliers, and every number it writes must exist in the data it was given.
const PROMPT_VERSION = "patterns-v1";

const SYSTEM = [
  "You are the pattern analyst inside UPAE, a payroll accuracy tool for a construction company.",
  "Code has already compared saved weekly hours (RT regular, OT 1.5x, DT 2x) and picked out statistical outliers. You never add outliers, never change them, and never calculate pay.",
  "You interpret ONLY the outliers supplied, for a payroll administrator. These are advisory observations, not rule violations and not accusations.",
  "",
  "Rules you must follow:",
  "1. Use ONLY the supplied data. Every number you write must appear exactly in the supplied data (facts, metrics, context). Do not compute new numbers, averages or percentages.",
  "2. A pattern can be completely legitimate (a standing Saturday crew, a long project push, a seasonal job). Say what makes it worth a look or what makes it look routine, using only the supplied facts. Never imply wrongdoing, never guess motives, never mention pay rates or dollars.",
  "3. The data has no job numbers or clock times at this level, so do not claim a cause. Offer at most one concrete check payroll can do in B2W or by asking the project manager.",
  "4. If there are only a few weeks of history, say so in dataLimits and lower your confidence.",
  "5. Use the key of each outlier exactly as supplied. Keep every field short and concrete.",
  "",
  "Reply with a single JSON object and nothing else, with exactly these keys:",
  '{"overview": string (1-2 sentences), "items": [{"key": string, "leaning": "worth_checking"|"likely_routine"|"unclear", "confidence": "low"|"medium"|"high", "interpretation": string (1-2 sentences), "evidence": string[] (1-3 short facts quoted from the supplied data), "suggestedCheck": string (one short sentence)}], "dataLimits": string (1-2 sentences)}',
].join("\n");

const s = (v, n) => String(v == null ? "" : v).slice(0, n || 200);
const num = (v) => { const x = Number(v); return isFinite(x) ? Math.round(x * 100) / 100 : 0; };
const list = (a, n) => (Array.isArray(a) ? a.slice(0, n) : []);

function trimInput(body) {
  if (!body || typeof body !== "object" || !Array.isArray(body.outliers) || !body.outliers.length) { const e = new Error("At least one outlier is required."); e.status = 400; throw e; }
  const metrics = (m) => { const o = {}; Object.keys(m || {}).slice(0, 12).forEach((k) => { const v = m[k]; if (typeof v === "number" && isFinite(v)) o[s(k, 30)] = num(v); else if (typeof v === "string") o[s(k, 30)] = s(v, 40); }); return o; };
  return {
    context: { weekStart: s(body.context && body.context.weekStart, 10), weekNo: body.context && body.context.weekNo != null ? num(body.context.weekNo) : null, weeksOfHistory: num(body.context && body.context.weeksOfHistory), priorWeeks: list(body.context && body.context.priorWeeks, 12).map((w) => s(w, 10)) },
    outliers: list(body.outliers, 15).map((o) => ({ key: s(o.key, 100), type: s(o.type, 30), employee: s(o.employee, 100), union: s(o.union, 80), facts: list(o.facts, 6).map((f) => s(f, 220)), metrics: metrics(o.metrics) })),
  };
}

function buildUserMessage(input) { return "Interpret these outliers. Data (JSON):\n" + JSON.stringify(input, null, 2); }

function numbersIn(text) { return (String(text).match(/\d+(?:,\d{3})*(?:\.\d+)?/g) || []).map((x) => String(Number(x.replace(/,/g, "")))); }
function checkNumbers(out, input) {
  const ok = new Set(numbersIn(JSON.stringify(input))), bad = [];
  const texts = [out.overview, out.dataLimits];
  out.items.forEach((i) => { texts.push(i.interpretation, i.suggestedCheck); i.evidence.forEach((e) => texts.push(e)); });
  texts.forEach((t) => numbersIn(t).forEach((n) => { if (!ok.has(n)) bad.push(n); }));
  return bad;
}

function parseOutput(text, input) {
  const m = String(text || "").match(/\{[\s\S]*\}/);
  if (!m) throw new Error("No JSON in reply");
  const o = JSON.parse(m[0]);
  if (typeof o.overview !== "string" || !o.overview.trim()) throw new Error("overview missing");
  const keys = new Set((input && input.outliers || []).map((x) => x.key)), seen = new Set();
  const items = (Array.isArray(o.items) ? o.items : []).filter((i) => i && keys.has(i.key) && !seen.has(i.key) && seen.add(i.key) && typeof i.interpretation === "string" && i.interpretation.trim()).map((i) => ({
    key: i.key,
    leaning: ["worth_checking", "likely_routine", "unclear"].includes(i.leaning) ? i.leaning : "unclear",
    confidence: ["low", "medium", "high"].includes(i.confidence) ? i.confidence : "low",
    interpretation: i.interpretation.slice(0, 400),
    evidence: (Array.isArray(i.evidence) ? i.evidence : []).filter((e) => typeof e === "string" && e.trim()).slice(0, 3).map((e) => e.slice(0, 250)),
    suggestedCheck: typeof i.suggestedCheck === "string" ? i.suggestedCheck.slice(0, 250) : "",
  }));
  if (!items.length) throw new Error("no usable items");
  const out = { overview: o.overview.slice(0, 500), items, dataLimits: typeof o.dataLimits === "string" ? o.dataLimits.slice(0, 500) : "" };
  if (input) { const bad = checkNumbers(out, input); if (bad.length) throw new Error("figures not in the data: " + bad.slice(0, 5).join(", ")); }
  return out;
}

async function callClaude(input, opts) {
  async function once(extra) {
    const res = await (opts.fetch || fetch)("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: { "content-type": "application/json", "x-api-key": opts.apiKey, "anthropic-version": "2023-06-01" },
      body: JSON.stringify({ model: opts.model, max_tokens: 3000, system: SYSTEM, messages: [{ role: "user", content: buildUserMessage(input) + (extra || "") }] }),
    });
    const body = await res.json().catch(() => ({}));
    if (opts.usage && body && body.usage) { opts.usage.inputTokens += Number(body.usage.input_tokens) || 0; opts.usage.outputTokens += Number(body.usage.output_tokens) || 0; opts.usage.calls += 1; }
    if (!res.ok) { const e = new Error("The AI service returned " + res.status + (body.error && body.error.message ? ": " + body.error.message : "")); e.status = res.status === 429 ? 429 : 502; throw e; }
    const txt = (body.content || []).map((c) => c.text || "").join("");
    if (body.stop_reason === "max_tokens") throw new Error("reply was cut off (too long)");
    return txt;
  }
  try { return parseOutput(await once(""), input); }
  catch (e1) {
    if (e1.status === 429 || (e1.status && e1.status >= 500)) throw e1;
    try { return parseOutput(await once("\n\nKeep every field short. Your previous reply was rejected (" + String(e1.message).slice(0, 120) + "). Reply with ONLY the JSON object, using only numbers that appear in the data."), input); }
    catch (e2) { const e = new Error("The pattern interpretation could not be checked against the data (" + String(e2.message).slice(0, 100) + "). Please try again."); e.status = 502; throw e; }
  }
}

function stubOutput(input) {
  return {
    overview: "[Sample answer: the AI key is not set up yet.] " + input.outliers.length + " outlier(s) were found by the statistical pass.",
    items: input.outliers.slice(0, 5).map((o) => ({ key: o.key, leaning: "unclear", confidence: "low", interpretation: "Sample text. Once the API key is added, this explains what makes the pattern worth a look.", evidence: o.facts.slice(0, 1), suggestedCheck: "Compare with B2W and ask the project manager if needed." })),
    dataLimits: "Sample answer.",
  };
}

module.exports = { PROMPT_VERSION, SYSTEM, trimInput, buildUserMessage, parseOutput, checkNumbers, callClaude, stubOutput };
