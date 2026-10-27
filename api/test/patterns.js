const assert = require("assert");
const crypto = require("crypto");
process.env.UPAE_TENANT_ID = "t-1"; process.env.UPAE_CLIENT_ID = "c-1"; delete process.env.ANTHROPIC_API_KEY;
const pt = require("../shared/patterns");
const body = { context: { weekStart: "2026-10-18", weekNo: 43, weeksOfHistory: 5, priorWeeks: ["2026-10-11", "2026-10-04"] }, secret: "x",
  outliers: [{ key: "pat|repeat_dt_day|AGUE1|Sat|2026-10-18", type: "repeat_dt_day", employee: "Aguiluz, Elvis (AGUE1)", union: "Local 14", facts: ["Double time on Saturday in 4 of 5 saved weeks (9 hours in total)."], metrics: { weeksWithPattern: 4, weeksObserved: 5, totalHours: 9, junk: { a: 1 } } },
    { key: "pat|union_premium|Local 14|2026-10-18", type: "union_premium", employee: "", union: "Local 14", facts: ["Overtime and double time were 18 percent of hours, against a typical 9 percent."], metrics: { sharePct: 18, typicalPct: 9 } }] };
const inp = pt.trimInput(body);
assert(!JSON.stringify(inp).includes("secret") && !JSON.stringify(inp).includes("junk"));
assert.throws(() => pt.trimInput({ outliers: [] }), /outlier/);
const good = { overview: "Two patterns in week 43 of 5 saved weeks.", items: [
  { key: body.outliers[0].key, leaning: "worth_checking", confidence: "medium", interpretation: "Double time on Saturday in 4 of 5 weeks looks like a standing pattern.", evidence: ["4 of 5 saved weeks", "9 hours in total"], suggestedCheck: "Ask the project manager if a Saturday crew is planned." },
  { key: "invented", leaning: "unclear", confidence: "low", interpretation: "x", evidence: [], suggestedCheck: "" },
  { key: body.outliers[1].key, leaning: "bogus", confidence: "bogus", interpretation: "Share rose from 9 to 18 percent.", evidence: ["18 percent of hours"], suggestedCheck: "" }], dataLimits: "Only 5 weeks of history." };
const out = pt.parseOutput("ok:\n" + JSON.stringify(good), inp);
assert.strictEqual(out.items.length, 2, "unknown keys dropped"); assert.strictEqual(out.items[1].leaning, "unclear"); assert.strictEqual(out.items[1].confidence, "low");
assert.throws(() => pt.parseOutput(JSON.stringify(Object.assign({}, good, { overview: "There were 31 cases." })), inp), /figures not in the data: 31/);
assert.throws(() => pt.parseOutput("nope", inp), /No JSON/);
const rep = (txt, i, o) => ({ ok: true, status: 200, json: async () => ({ content: [{ type: "text", text: txt }], usage: { input_tokens: i, output_tokens: o } }) });
(async () => {
  const usage = { inputTokens: 0, outputTokens: 0, calls: 0 }; const seq = [rep("bad", 500, 10), rep(JSON.stringify(good), 600, 200)]; let n = 0;
  const r = await pt.callClaude(inp, { apiKey: "k", model: "m", usage, fetch: async (u, o) => { assert(JSON.parse(o.body).system.includes("outliers")); return seq[n++]; } });
  assert.strictEqual(r.items.length, 2); assert.deepStrictEqual(usage, { inputTokens: 1100, outputTokens: 210, calls: 2 });
  await assert.rejects(() => pt.callClaude(inp, { apiKey: "k", model: "m", fetch: async () => rep("x", 1, 1) }), (e) => e.status === 502);
  assert(pt.stubOutput(inp).overview.includes("Sample"));
  const { publicKey, privateKey } = crypto.generateKeyPairSync("rsa", { modulusLength: 2048 });
  const jwk = Object.assign(publicKey.export({ format: "jwk" }), { kid: "k", alg: "RS256" });
  global.fetch = async () => ({ ok: true, json: async () => ({ keys: [jwk] }) });
  const b64 = (o) => Buffer.from(JSON.stringify(o)).toString("base64url");
  const tok = (email) => { const h = b64({ alg: "RS256", kid: "k" }), t = Math.floor(Date.now() / 1000), p = b64({ aud: "c-1", tid: "t-1", iss: "https://login.microsoftonline.com/t-1/v2.0", exp: t + 600, preferred_username: email });
    return h + "." + p + "." + crypto.sign("RSA-SHA256", Buffer.from(h + "." + p), privateKey).toString("base64url"); };
  const handler = require("../patterns/index.js");
  const run = async (headers, b) => { const ctx = { log: Object.assign(() => {}, { error() {} }) }; await handler(ctx, { headers, body: b }); return ctx.res; };
  assert.strictEqual((await run({}, body)).status, 401);
  const res = await run({ "x-upae-token": "Bearer " + tok("a@b.com") }, body);
  assert.strictEqual(res.status, 200); assert.strictEqual(res.body.stub, true); assert.strictEqual(res.body.promptVersion, "patterns-v1");
  assert.strictEqual((await run({ "x-upae-token": "Bearer " + tok("a@b.com") }, { outliers: [] })).status, 400);
  console.log("patterns tests passed");
})().catch((e) => { console.error(e); process.exit(1); });
