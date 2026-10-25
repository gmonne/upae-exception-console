// Prompt, input trimming, output validation and the Claude call for "Explain a finding" (Phase 1).
const PROMPT_VERSION = "explain-v1";

const SYSTEM = [
  "You are the explanation assistant inside UPAE, a payroll accuracy tool for a construction company.",
  "A deterministic rules engine has already decided whether an employee's hours for a day match the union rule. You never change that decision, and you never calculate pay.",
  "Your job is to explain, in plain language for a payroll administrator, why this record was flagged, what most likely happened, and what to check or correct.",
  "",
  "Rules you must follow:",
  "1. Use ONLY the rule text, record and week pattern given in the user message. Do not use outside knowledge of union agreements.",
  "2. Hours are reported as RT (regular), OT (overtime at 1.5x) and DT (double time at 2x). Do not invent other categories.",
  "3. If the supplied information is not enough to say why, set needsMoreInfo to true and say what is missing. Never guess.",
  "4. The engine's expected split is correct by definition; do not argue with it. If it looks inconsistent with the rule text, say so under needsMoreInfo.",
  "5. Do not mention pay rates or dollar amounts. Do not give legal advice.",
  "6. Keep every field short and concrete. Evidence must quote or closely paraphrase lines of the supplied rule text.",
  "",
  "Reply with a single JSON object and nothing else, with exactly these keys:",
  '{"explanation": string (1-3 sentences), "likelyCause": string (1-2 sentences), "suggestedAction": string (1-2 sentences, what the reviewer or project manager should check or correct), "confidence": "low"|"medium"|"high", "evidence": string[] (1-4 items drawn from the rule text), "needsMoreInfo": boolean}',
].join("\n");

const s = (v, n) => String(v == null ? "" : v).slice(0, n || 400);

// Keep only the fields the task needs (data minimisation). Throws on bad input.
function trimInput(body) {
  if (!body || typeof body !== "object" || !body.finding) { const e = new Error("A finding is required."); e.status = 400; throw e; }
  const f = body.finding, r = body.rule || null;
  const fields = {};
  if (r && r.fields && typeof r.fields === "object") {
    Object.keys(r.fields).slice(0, 40).forEach((k) => { const v = s(r.fields[k], 600); if (v) fields[s(k, 60)] = v; });
  }
  return {
    week: { weekNo: body.week && body.week.weekNo || null, start: s(body.week && body.week.start, 10), end: s(body.week && body.week.end, 10) },
    finding: {
      kind: s(f.kind, 30), employeeId: s(f.id, 20), employeeName: s(f.name, 80), union: s(f.union, 80), day: s(f.iso, 10), dayName: s(f.dayName, 12),
      reported: s(f.reported, 60), engineExpected: s(f.expected, 60), engineMessage: s(f.msg, 400), engineReasons: (Array.isArray(f.broken) ? f.broken : []).slice(0, 6).map((x) => s(x, 400)),
    },
    rule: r ? { id: s(r.id, 20), name: s(r.name, 100), cba: s(r.cba, 150), fields } : null,
    weekPattern: (Array.isArray(body.weekPattern) ? body.weekPattern : []).slice(0, 7).map((d) => ({ day: s(d.day, 12), rt: +d.rt || 0, ot: +d.ot || 0, dt: +d.dt || 0 })),
  };
}

function buildUserMessage(input) {
  return "Explain this flagged record. Data (JSON):\n" + JSON.stringify(input, null, 2);
}

// Parse and validate the model's reply. Returns an object or throws.
function parseOutput(text) {
  let t = String(text || "").trim();
  const m = t.match(/\{[\s\S]*\}/);
  if (!m) throw new Error("No JSON in reply");
  const o = JSON.parse(m[0]);
  const conf = ["low", "medium", "high"].includes(o.confidence) ? o.confidence : null;
  if (typeof o.explanation !== "string" || !o.explanation.trim()) throw new Error("explanation missing");
  if (typeof o.likelyCause !== "string" || typeof o.suggestedAction !== "string") throw new Error("fields missing");
  if (!conf) throw new Error("confidence invalid");
  const ev = Array.isArray(o.evidence) ? o.evidence.filter((x) => typeof x === "string").slice(0, 4).map((x) => x.slice(0, 400)) : [];
  return {
    explanation: o.explanation.slice(0, 800), likelyCause: o.likelyCause.slice(0, 600), suggestedAction: o.suggestedAction.slice(0, 600),
    confidence: conf, evidence: ev, needsMoreInfo: !!o.needsMoreInfo,
  };
}

async function callClaude(input, opts) {
  const key = opts.apiKey, model = opts.model;
  async function once(extra) {
    const res = await (opts.fetch || fetch)("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: { "content-type": "application/json", "x-api-key": key, "anthropic-version": "2023-06-01" },
      body: JSON.stringify({ model, max_tokens: 900, system: SYSTEM, messages: [{ role: "user", content: buildUserMessage(input) + (extra || "") }] }),
    });
    const body = await res.json().catch(() => ({}));
    if (!res.ok) { const e = new Error("The AI service returned " + res.status + (body.error && body.error.message ? ": " + body.error.message : "")); e.status = res.status === 429 ? 429 : 502; throw e; }
    return (body.content || []).map((c) => c.text || "").join("");
  }
  let text = await once("");
  try { return parseOutput(text); }
  catch (e1) {
    text = await once("\n\nYour previous reply was not valid. Reply with ONLY the JSON object described.");
    return parseOutput(text);
  }
}

// Used when no API key is configured yet, so the screen can be tested end to end.
function stubOutput(input) {
  const f = input.finding;
  return {
    explanation: "[Sample answer: the AI key is not set up yet.] " + (f.engineMessage || "This record was flagged by the rules engine."),
    likelyCause: "Sample text. Once the API key is added, this will describe the most likely cause using the rule text.",
    suggestedAction: "Review the record against the rule shown, then correct it in B2W if needed.",
    confidence: "low", evidence: input.rule ? ["(sample) " + (input.rule.name || "rule")] : [], needsMoreInfo: true,
  };
}

module.exports = { PROMPT_VERSION, SYSTEM, trimInput, buildUserMessage, parseOutput, callClaude, stubOutput };
