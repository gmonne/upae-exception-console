const { verifyToken } = require("../shared/auth");
const { makeGraph } = require("../shared/graph");
const R = require("../shared/review");

const hits = new Map();
const LIMIT = parseInt(process.env.UPAE_REVIEW_RATE_LIMIT || "120", 10), WINDOW_MS = 10 * 60 * 1000;
function limited(user) {
  const now = Date.now(), arr = (hits.get(user) || []).filter((t) => now - t < WINDOW_MS);
  if (arr.length >= LIMIT) { hits.set(user, arr); return true; }
  arr.push(now); hits.set(user, arr); return false;
}
const list = (v) => String(v || "").toLowerCase().split(/[;,\s]+/).filter(Boolean);

// Factory so tests can inject a fake Graph and token check.
function makeHandler(deps) {
  const d = Object.assign({ verify: verifyToken, graph: null, now: () => new Date() }, deps || {});
  return async function (context, req) {
    const send = (status, body) => { context.res = { status, headers: { "Content-Type": "application/json", "Cache-Control": "no-store" }, body }; };
    try {
      const user = await d.verify(req.headers["x-upae-token"] || req.headers["authorization"]);
      if (!user.email) return send(401, { error: "Your sign-in did not include an email address." });
      if (limited(user.email)) return send(429, { error: "Too many requests. Please wait a few minutes." });
      const raw = typeof req.body === "string" ? JSON.parse(req.body || "{}") : (req.body || {});
      if (JSON.stringify(raw).length > 60000) return send(413, { error: "Request too large." });
      const action = (context.bindingData && context.bindingData.action) || "";
      const graph = d.graph || (d.graph = makeGraph());

      // A named administrator may open the page as a project manager (to check what they see).
      let pm = user.email, actingAs = "";
      if (raw.as) {
        if (!list(process.env.UPAE_REVIEW_ADMINS).includes(user.email)) return send(403, { error: "You are not allowed to open the page for someone else." });
        pm = String(raw.as).toLowerCase().trim(); actingAs = pm;
      }
      const mine = R.pmJobs(await graph.listItems(R.LISTS.jobs), pm);

      if (action === "list") {
        const items = R.selectFor(await graph.listItems(R.LISTS.findings), mine);
        return send(200, { pm, actingAs, jobs: Object.keys(mine).sort().map((k) => mine[k]), items });
      }
      if (action === "respond") {
        const batch = (Array.isArray(raw.items) ? raw.items : [raw]).slice(0, 60);
        if (!batch.length) return send(400, { error: "Nothing to save." });
        const who = actingAs ? user.email + " (viewing as " + actingAs + ")" : user.email;
        const stamp = d.now().toISOString().replace("T", " ").slice(0, 16) + " UTC";
        const results = [];
        for (const b of batch) {
          try {
            const v = R.validateResponse(b);
            const it = await graph.getItem(R.LISTS.findings, b.id);
            if (!R.canSee(it, mine)) throw Object.assign(new Error("This record is not one of your jobs, or it is no longer waiting for a response."), { status: 403 });
            await graph.patchFields(R.LISTS.findings, b.id, { PMResponse: v.response, PMNote: v.note, PMRespondedBy: who, PMRespondedDate: stamp });
            const f = it.fields || {};
            try {
              await graph.createItem(R.LISTS.history, { Title: String(f.FindingKey || "").slice(0, 250), FindingKey: f.FindingKey || "", WeekStart: f.WeekStart || "", EmployeeID: f.EmployeeID || "", EmployeeName: f.EmployeeName || "",
                FromStatus: f.PMResponse ? "PM: " + f.PMResponse : "Sent to PM", ToStatus: "PM: " + v.response, Detail: v.note, ChangedBy: who });
            } catch (e) { /* history is best effort */ }
            results.push({ id: b.id, ok: true, response: v.response, note: v.note, by: who, date: stamp });
          } catch (e) { results.push({ id: b.id, ok: false, error: e.message || "Failed", status: e.status || 500 }); }
        }
        return send(200, { results });
      }
      return send(404, { error: "Unknown action." });
    } catch (err) {
      const status = err.status || 500;
      if (status >= 500) context.log && context.log.error && context.log.error("review failed: " + (err && err.message));
      return send(status, { error: status >= 500 && status !== 502 && status !== 503 ? "The review service had a problem. Please try again." : (err.message || "Request failed.") });
    }
  };
}
module.exports = makeHandler();
module.exports.makeHandler = makeHandler;
