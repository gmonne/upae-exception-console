const assert = require("assert");
process.env.UPAE_REVIEW_ADMINS = "gus@cacindinc.com";
const { makeHandler } = require("../review/index");
const R = require("../shared/review");

const J = [
  { id: "j1", fields: { Title: "23-604", JobName: "Rosedale", PMEmail: "tchou@cacindinc.com", PMName: "Tai Chou", Active: true } },
  { id: "j2", fields: { Title: "24-110", PMEmail: "erogan@cacindinc.com", Active: true } },
  { id: "j3", fields: { Title: "25-615", PMEmail: "tchou@cacindinc.com", Active: false } },
];
const mk = (id, over) => ({ id, fields: Object.assign({ FindingKey: "k" + id, WeekStart: "2026-10-18", WeekNo: 43, EmployeeID: "AGUE1", EmployeeName: "Aguiluz, Elvis", WorkDate: "2026-10-19", Kind: "mismatch", Status: "Sent to PM",
  Jobs: "23-604~0099 (13.5 OT); 24-110 (8 RT + 1 OT)", Reported: "8 RT + 14.5 OT", Expected: "8 RT + 7 OT + 7.5 DT", AIExplanation: "x" }, over || {}) });
let F;
const patched = [], created = [];
const graph = {
  listItems: async (l) => (l === "JobAssignments" ? J : F),
  getItem: async (l, id) => F.find((x) => x.id === id),
  patchFields: async (l, id, f) => { patched.push([id, f]); Object.assign(F.find((x) => x.id === id).fields, f); },
  createItem: async (l, f) => { created.push([l, f]); },
};
const call = async (email, action, body) => {
  const h = makeHandler({ verify: async () => ({ email }), graph, now: () => new Date("2026-10-27T14:05:00Z") });
  const ctx = { bindingData: { action }, log: { error() {} } }; await h(ctx, { headers: { "x-upae-token": "Bearer t" }, body }); return ctx.res;
};
(async () => {
  F = [mk("1"), mk("2", { Jobs: "24-110 (8 RT)" }), mk("3", { Status: "Resolved" }), mk("4", { Status: "Open" }), mk("5", { Jobs: "99-999 (8 RT)" }), mk("6", { Jobs: "25-615 (8 RT)" })];
  // jobs / bases
  assert.strictEqual(R.jobBase("23-430~0328"), "23-430"); assert.strictEqual(R.jobBase("25-615-0103"), "25-615"); assert.strictEqual(R.jobBase("15-800"), "15-800");
  // Tai sees only Sent-to-PM records on 23-604
  let r = await call("tchou@cacindinc.com", "list", {});
  assert.strictEqual(r.status, 200); assert.deepStrictEqual(r.body.items.map((i) => i.id), ["1"]); assert.deepStrictEqual(r.body.jobs.map((j) => j.job), ["23-604"]);
  assert.deepStrictEqual(r.body.items[0].myJobs, ["23-604"]); assert.ok(r.body.items[0].jobs.includes("24-110"));   // other job shown as context
  // Eric sees the two-job day too, and his own record
  r = await call("erogan@cacindinc.com", "list", {}); assert.deepStrictEqual(r.body.items.map((i) => i.id).sort(), ["1", "2"]);
  // someone who owns no job sees nothing
  r = await call("nobody@cacindinc.com", "list", {}); assert.strictEqual(r.body.items.length, 0);
  // `as` is for administrators only
  r = await call("tchou@cacindinc.com", "list", { as: "erogan@cacindinc.com" }); assert.strictEqual(r.status, 403);
  r = await call("gus@cacindinc.com", "list", { as: "tchou@cacindinc.com" }); assert.strictEqual(r.status, 200); assert.strictEqual(r.body.items.length, 1); assert.strictEqual(r.body.actingAs, "tchou@cacindinc.com");
  // respond: validation
  r = await call("tchou@cacindinc.com", "respond", { id: "1", response: "Needs correction", note: "" }); assert.strictEqual(r.body.results[0].ok, false);
  r = await call("tchou@cacindinc.com", "respond", { id: "1", response: "Bogus" }); assert.strictEqual(r.body.results[0].ok, false);
  // not his record / already resolved / not yet sent
  r = await call("tchou@cacindinc.com", "respond", { items: [{ id: "2", response: "Correct as reported" }, { id: "3", response: "Correct as reported" }, { id: "4", response: "Correct as reported" }, { id: "5", response: "Correct as reported" }] });
  assert.ok(r.body.results.every((x) => !x.ok && x.status === 403)); assert.strictEqual(patched.length, 0);
  // good response
  r = await call("tchou@cacindinc.com", "respond", { id: "1", response: "Needs correction", note: "Move 7 OT to DT" });
  assert.ok(r.body.results[0].ok); assert.deepStrictEqual(patched[0], ["1", { PMResponse: "Needs correction", PMNote: "Move 7 OT to DT", PMRespondedBy: "tchou@cacindinc.com", PMRespondedDate: "2026-10-27 14:05 UTC" }]);
  assert.strictEqual(created.length, 1); assert.strictEqual(created[0][1].ToStatus, "PM: Needs correction"); assert.strictEqual(created[0][1].ChangedBy, "tchou@cacindinc.com");
  // the other PM on the same record can answer too; latest wins, shown to Payroll
  r = await call("erogan@cacindinc.com", "respond", { id: "1", response: "Correct as reported" }); assert.ok(r.body.results[0].ok);
  // admin writes are labelled
  r = await call("gus@cacindinc.com", "respond", { as: "tchou@cacindinc.com", id: "1", response: "Question", note: "why?" }); assert.ok(r.body.results[0].by.includes("viewing as tchou@cacindinc.com"));
  // after Payroll resolves it, a PM can no longer change it
  F[0].fields.Status = "Resolved"; r = await call("tchou@cacindinc.com", "respond", { id: "1", response: "Correct as reported" }); assert.strictEqual(r.body.results[0].ok, false);
  // setup missing -> clear message
  const h = makeHandler({ verify: async () => ({ email: "tchou@cacindinc.com" }) }); const ctx = { bindingData: { action: "list" }, log: {} };
  await h(ctx, { headers: { "x-upae-token": "Bearer t" }, body: {} }); assert.strictEqual(ctx.res.status, 503); assert.ok(/not set up/.test(ctx.res.body.error));
  // unauthenticated
  const h2 = makeHandler({ verify: async () => { throw Object.assign(new Error("Missing sign-in token."), { status: 401 }); }, graph }); const c2 = { bindingData: { action: "list" }, log: {} };
  await h2(c2, { headers: {}, body: {} }); assert.strictEqual(c2.res.status, 401);
  console.log("review tests passed");
})().catch((e) => { console.error(e); process.exit(1); });
