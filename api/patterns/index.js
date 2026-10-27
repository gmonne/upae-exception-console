const { verifyToken } = require("../shared/auth");
const { PROMPT_VERSION, trimInput, callClaude, stubOutput } = require("../shared/patterns");

// Best-effort per-user limit (resets when the instance recycles). The real cap is the spend limit on the API key.
const hits = new Map();
const LIMIT = parseInt(process.env.UPAE_AI_PATTERNS_RATE_LIMIT || "8", 10), WINDOW_MS = 10 * 60 * 1000;
function limited(user) {
  const now = Date.now(), arr = (hits.get(user) || []).filter((t) => now - t < WINDOW_MS);
  if (arr.length >= LIMIT) { hits.set(user, arr); return true; }
  arr.push(now); hits.set(user, arr); return false;
}

module.exports = async function (context, req) {
  const send = (status, body) => { context.res = { status, headers: { "Content-Type": "application/json", "Cache-Control": "no-store" }, body }; };
  try {
    const user = await verifyToken(req.headers["x-upae-token"] || req.headers["authorization"]);
    const allow = (process.env.UPAE_AI_ALLOWED_USERS || "").toLowerCase().split(/[;,\s]+/).filter(Boolean);
    if (allow.length && !allow.includes(user.email)) return send(403, { error: "You are not on the list of people who can use the AI features." });
    if (limited(user.email || user.oid)) return send(429, { error: "Too many pattern requests in a short time. Please wait a few minutes." });

    const raw = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
    if (JSON.stringify(raw || {}).length > 40000) return send(413, { error: "Request too large." });
    const input = trimInput(raw);

    const apiKey = process.env.ANTHROPIC_API_KEY;
    const model = process.env.ANTHROPIC_MODEL || "claude-sonnet-5-5";
    if (!apiKey || process.env.UPAE_AI_STUB === "1") {
      return send(200, Object.assign({ stub: true, model: "stub", promptVersion: PROMPT_VERSION }, stubOutput(input)));
    }
    const usage = { inputTokens: 0, outputTokens: 0, calls: 0 };
    try {
      const out = await callClaude(input, { apiKey, model, usage });
      return send(200, Object.assign({ stub: false, model, promptVersion: PROMPT_VERSION, usage }, out));
    } catch (e) { e.usage = usage; throw e; }
  } catch (err) {
    const status = err.status || 500;
    if (status >= 500) context.log.error("patterns failed: " + (err && err.message));
    return send(status, { error: status >= 500 && status !== 502 ? "The AI request failed. Please try again." : (err.message || "Request failed."), usage: err.usage && err.usage.calls ? err.usage : undefined });
  }
};
