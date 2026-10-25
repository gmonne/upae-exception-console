const assert = require("assert");
const crypto = require("crypto");
process.env.UPAE_TENANT_ID = "t-1"; process.env.UPAE_CLIENT_ID = "c-1"; delete process.env.ANTHROPIC_API_KEY;
const { publicKey, privateKey } = crypto.generateKeyPairSync("rsa", { modulusLength: 2048 });
const jwk = Object.assign(publicKey.export({ format: "jwk" }), { kid: "k", alg: "RS256" });
global.fetch = async (url) => { assert(/login\.microsoftonline\.com/.test(url)); return { ok: true, json: async () => ({ keys: [jwk] }) }; };
const b64 = (o) => Buffer.from(JSON.stringify(o)).toString("base64url");
function tok(email) { const h = b64({ alg: "RS256", kid: "k" }), n = Math.floor(Date.now() / 1000), p = b64({ aud: "c-1", tid: "t-1", iss: "https://login.microsoftonline.com/t-1/v2.0", exp: n + 600, preferred_username: email });
  return h + "." + p + "." + crypto.sign("RSA-SHA256", Buffer.from(h + "." + p), privateKey).toString("base64url"); }
const handler = require("../explain/index.js");
const run = async (headers, body) => { const ctx = { log: Object.assign(() => {}, { error() {} }) }; await handler(ctx, { headers, body }); return ctx.res; };
(async () => {
  const body = { finding: { kind: "mismatch", id: "A1", msg: "Reported 8 RT + 2 OT" }, rule: { name: "R" } };
  assert.strictEqual((await run({}, body)).status, 401);
  assert.strictEqual((await run({ authorization: "Bearer eyJhbGciOiJIUzI1NiJ9.e30.x" }, body)).status, 401);
  let r = await run({ "x-upae-token": "Bearer " + tok("a@b.com"), authorization: "Bearer eyJhbGciOiJIUzI1NiJ9.e30.x" }, body);
  assert.strictEqual(r.status, 200); assert.strictEqual(r.body.stub, true); assert(r.body.explanation.includes("Sample"));
  assert.strictEqual((await run({ authorization: "Bearer " + tok("a@b.com") }, {})).status, 400);
  process.env.UPAE_AI_ALLOWED_USERS = "x@y.com";
  assert.strictEqual((await run({ authorization: "Bearer " + tok("a@b.com") }, body)).status, 403);
  delete process.env.UPAE_AI_ALLOWED_USERS; process.env.UPAE_AI_RATE_LIMIT = "1";
  console.log("handler tests passed");
})().catch((e) => { console.error("FAIL", e); process.exit(1); });
