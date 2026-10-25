const assert = require("assert");
const crypto = require("crypto");
const { verifyToken } = require("../shared/auth");
const ex = require("../shared/explain");

const TENANT = "t-123", CLIENT = "c-456";
const { publicKey, privateKey } = crypto.generateKeyPairSync("rsa", { modulusLength: 2048 });
const jwk = Object.assign(publicKey.export({ format: "jwk" }), { kid: "k1", alg: "RS256", use: "sig" });
const b64 = (o) => Buffer.from(typeof o === "string" ? o : JSON.stringify(o)).toString("base64url");
function token(over, kid) {
  const h = b64({ alg: "RS256", kid: kid || "k1", typ: "JWT" });
  const now = Math.floor(Date.now() / 1000);
  const p = b64(Object.assign({ aud: CLIENT, tid: TENANT, iss: "https://login.microsoftonline.com/" + TENANT + "/v2.0", exp: now + 600, nbf: now - 5, preferred_username: "Gus@Example.com", name: "Gus" }, over));
  const sig = crypto.sign("RSA-SHA256", Buffer.from(h + "." + p), privateKey).toString("base64url");
  return h + "." + p + "." + sig;
}
const opts = { tenantId: TENANT, clientId: CLIENT, getKeys: async () => [jwk] };
const rej = async (promise, re, label) => { try { await promise; } catch (e) { assert(re.test(e.message), label + ": " + e.message); assert.strictEqual(e.status, 401); return; } assert.fail(label + " should have failed"); };

(async () => {
  const u = await verifyToken("Bearer " + token(), opts);
  assert.strictEqual(u.email, "gus@example.com");
  await rej(verifyToken(null, opts), /Missing/, "no header");
  await rej(verifyToken("Bearer " + token({ exp: 1 }), opts), /expired/, "expired");
  await rej(verifyToken("Bearer " + token({ aud: "other" }), opts), /different application/, "aud");
  await rej(verifyToken("Bearer " + token({ tid: "x" }), opts), /different organization/, "tid");
  await rej(verifyToken("Bearer " + token({ iss: "https://evil.example/" + TENANT }), opts), /issuer/, "iss");
  await rej(verifyToken("Bearer " + token({}, "nokid"), opts), /Unknown signing key/, "kid");
  const t = token(); const bad = t.slice(0, -4) + (t.slice(-4) === "AAAA" ? "BBBB" : "AAAA");
  await rej(verifyToken("Bearer " + bad, opts), /signature/, "tamper");
  console.log("auth ok");

  const body = { week: { weekNo: 43, start: "2026-10-18", end: "2026-10-24" }, finding: { kind: "mismatch", id: "DONR1", name: "Donnellan", union: "Local 14", iso: "2026-10-19", reported: "8 RT + 2 OT", expected: "8 RT + 2 DT", msg: "m", broken: ["no 1.5x tier"], secretPay: "$99/hr" },
    rule: { id: "R-001", name: "Local 14", cba: "Art. 12", fields: { "OT 2x threshold": "Over 8 hrs/day", "Hourly Rate": "" } }, weekPattern: [{ day: "Mon", rt: 8, ot: 2, dt: 0 }], extra: "ignored" };
  const inp = ex.trimInput(body);
  assert(!JSON.stringify(inp).includes("99/hr") && !JSON.stringify(inp).includes("ignored"), "unexpected data passed through");
  assert.strictEqual(inp.rule.fields["Hourly Rate"], undefined);
  assert.throws(() => ex.trimInput({}), /finding is required/);
  const good = { explanation: "e", likelyCause: "c", suggestedAction: "a", confidence: "high", evidence: ["x"], needsMoreInfo: false };
  assert.deepStrictEqual(ex.parseOutput("Sure!\n" + JSON.stringify(good)).confidence, "high");
  assert.throws(() => ex.parseOutput("nope"), /No JSON/);
  assert.throws(() => ex.parseOutput(JSON.stringify(Object.assign({}, good, { confidence: "certain" }))), /confidence/);

  let calls = 0;
  const mk = (replies) => async (url, o) => { const r = replies[calls++]; const sent = JSON.parse(o.body); assert.strictEqual(sent.temperature, 0); assert(sent.system.includes("Never guess") || sent.system.includes("never guess") || sent.system.includes("Never")); return { ok: r.ok !== false, status: r.status || 200, json: async () => r.body }; };
  const reply = (txt) => ({ body: { content: [{ type: "text", text: txt }] } });
  let r = await ex.callClaude(inp, { apiKey: "k", model: "m", fetch: mk([reply(JSON.stringify(good))]) });
  assert.strictEqual(calls, 1); assert.strictEqual(r.explanation, "e");
  calls = 0; r = await ex.callClaude(inp, { apiKey: "k", model: "m", fetch: mk([reply("garbage"), reply(JSON.stringify(good))]) });
  assert.strictEqual(calls, 2);
  calls = 0; await assert.rejects(() => ex.callClaude(inp, { apiKey: "k", model: "m", fetch: mk([{ ok: false, status: 429, body: { error: { message: "slow down" } } }]) }), (e) => e.status === 429);
  console.log("explain ok");

  // handler
  const handler = require("../explain/index.js");
  process.env.UPAE_TENANT_ID = TENANT; process.env.UPAE_CLIENT_ID = CLIENT;
  console.log("all tests passed");
})().catch((e) => { console.error("FAIL", e); process.exit(1); });
