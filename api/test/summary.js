const assert = require("assert");
const crypto = require("crypto");
process.env.UPAE_TENANT_ID = "t-1"; process.env.UPAE_CLIENT_ID = "c-1"; delete process.env.ANTHROPIC_API_KEY;
const sm = require("../shared/summary");

const body = { week: { weekNo: 43, start: "2026-10-18", end: "2026-10-24" }, totals: { employeesInFile: 441, rt: 9000.5, ot: 700, dt: 300, totalHours: 10000.5, otSharePct: 7, dtSharePct: 3 },
  coverage: { employeesChecked: 120, daysChecked: 500, daysMatching: 480, checkedHoursPct: 40, unionsSkipped: [{ union: "Local 9", employees: 12 }] },
  findings: { total: 25, byKind: [{ kind: "mismatch", count: 20 }], byStatus: { open: 20, sentToPm: 3, resolved: 2 }, ruleDifferencesByUnion: [{ union: "Local 14", differences: 12, daysChecked: 60 }],
    employeesWithMostDifferences: [{ employee: "Aguiluz, Elvis (AGUE1)", differences: 4 }], noRtDayDifferences: 5, resolvedReasons: [{ reason: "Keying error", count: 2 }], secret: "x" },
  priorWeeks: [{ weekNo: 42, start: "2026-10-11", totalHours: 9800, ruleDifferences: 18, otSharePct: 6 }],
  candidates: [{ key: "2026-10-18|AGUE1|2026-10-19|mismatch", employee: "Aguiluz", union: "Local 1010", day: "Mon 10/19", reported: "13.5 OT", ruleSays: "8 RT + 5.5 OT", kind: "mismatch", status: "Open", noRtDay: true, hasAiNote: true, repeatEmployeeCount: 4, payRate: "$50" }], extra: "ignored" };
const inp = sm.trimInput(body);
assert(!JSON.stringify(inp).includes("ignored") && !JSON.stringify(inp).includes("$50") && !JSON.stringify(inp).includes('"secret"'), "unexpected data passed through");
assert.throws(() => sm.trimInput({}), /week with totals/);

const good = { headline: "Week 43 had 25 findings, 20 still open.", keyChanges: ["Rule differences rose from 18 to 20 compared with week 42."], riskRanked: [{ key: body.candidates[0].key, risk: "high", reason: "Day with 13.5 OT and no RT; verify second shift." }, { key: "unknown|key", risk: "low", reason: "dropped" }],
  suggestedFollowUps: ["Email the project managers about the 20 open items."], dataLimits: "No job numbers on the report.", confidence: "medium" };
const out = sm.parseOutput("Here:\n" + JSON.stringify(good), inp);
assert.strictEqual(out.riskRanked.length, 1, "unknown keys must be dropped");
// invented figure is rejected
assert.throws(() => sm.parseOutput(JSON.stringify(Object.assign({}, good, { headline: "Week 43 had 31 findings." })), inp), /figures not in the data: 31/);
// dates and comma numbers from the data are allowed
sm.parseOutput(JSON.stringify(Object.assign({}, good, { headline: "Total of 10,000.5 hours in week 43 (10/18 to 10/24)." })), inp);
assert.throws(() => sm.parseOutput("nope", inp), /No JSON/);
// job focus: only jobs in byJob, numbers checked
const jb = sm.trimInput(Object.assign({}, body, { source: "detail", byJob: [{ job: "20-103", name: "Manhattan Gas", pm: "Gus Monne", findings: 33, open: 20, sentToPm: 13, pmAnswered: 2, resolved: 0 }], pmWaiting: [{ pm: "Gus Monne", waiting: 11, answered: 2 }], jobsWithoutPm: 4 }));
const withJobs = Object.assign({}, good, { jobFocus: [{ job: "20-103", note: "33 findings, 11 waiting on Gus Monne." }, { job: "99-999", note: "invented job" }] });
const oj = sm.parseOutput(JSON.stringify(withJobs), jb);
assert.strictEqual(oj.jobFocus.length, 1, "unknown jobs must be dropped");
assert.throws(() => sm.parseOutput(JSON.stringify(Object.assign({}, good, { jobFocus: [{ job: "20-103", note: "47 findings." }] })), jb), /figures not in the data: 47/);
assert.strictEqual(sm.trimInput(body).source, "summary");

let calls = 0;
const reply = (txt) => ({ ok: true, status: 200, json: async () => ({ content: [{ type: "text", text: txt }] }) });
const mk = (replies) => async (url, o) => { const sent = JSON.parse(o.body); assert(sent.system.includes("ONE week")); return replies[calls++]; };
(async () => {
  let r = await sm.callClaude(inp, { apiKey: "k", model: "m", fetch: mk([reply(JSON.stringify(good))]) });
  assert.strictEqual(calls, 1); assert.strictEqual(r.confidence, "medium");
  calls = 0; r = await sm.callClaude(inp, { apiKey: "k", model: "m", fetch: mk([reply(JSON.stringify(Object.assign({}, good, { headline: "Made up 77 findings" }))), reply(JSON.stringify(good))]) });
  assert.strictEqual(calls, 2);
  calls = 0; await assert.rejects(() => sm.callClaude(inp, { apiKey: "k", model: "m", fetch: mk([reply("x"), reply("y")]) }), (e) => e.status === 502 && /could not be checked/.test(e.message));
  calls = 0; await assert.rejects(() => sm.callClaude(inp, { apiKey: "k", model: "m", fetch: mk([{ ok: false, status: 429, json: async () => ({ error: { message: "slow" } }) }]) }), (e) => e.status === 429);
  assert(sm.stubOutput(inp).headline.includes("Sample"));

  // handler
  const { publicKey, privateKey } = crypto.generateKeyPairSync("rsa", { modulusLength: 2048 });
  const jwk = Object.assign(publicKey.export({ format: "jwk" }), { kid: "k", alg: "RS256" });
  global.fetch = async (url) => { assert(/login\.microsoftonline\.com/.test(url)); return { ok: true, json: async () => ({ keys: [jwk] }) }; };
  const b64 = (o) => Buffer.from(JSON.stringify(o)).toString("base64url");
  const tok = (email) => { const h = b64({ alg: "RS256", kid: "k" }), n = Math.floor(Date.now() / 1000), p = b64({ aud: "c-1", tid: "t-1", iss: "https://login.microsoftonline.com/t-1/v2.0", exp: n + 600, preferred_username: email });
    return h + "." + p + "." + crypto.sign("RSA-SHA256", Buffer.from(h + "." + p), privateKey).toString("base64url"); };
  const handler = require("../summary/index.js");
  const run = async (headers, b) => { const ctx = { log: Object.assign(() => {}, { error() {} }) }; await handler(ctx, { headers, body: b }); return ctx.res; };
  assert.strictEqual((await run({}, body)).status, 401);
  let res = await run({ "x-upae-token": "Bearer " + tok("a@b.com"), authorization: "Bearer eyJhbGciOiJIUzI1NiJ9.e30.x" }, body);
  assert.strictEqual(res.status, 200); assert.strictEqual(res.body.stub, true); assert(res.body.headline.includes("Sample"));
  assert.strictEqual((await run({ "x-upae-token": "Bearer " + tok("a@b.com") }, {})).status, 400);
  process.env.UPAE_AI_ALLOWED_USERS = "x@y.com";
  assert.strictEqual((await run({ "x-upae-token": "Bearer " + tok("a@b.com") }, body)).status, 403);
  console.log("summary tests passed");
})().catch((e) => { console.error("FAIL", e); process.exit(1); });
