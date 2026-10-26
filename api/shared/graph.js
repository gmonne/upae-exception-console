// Microsoft Graph with the app's own identity (client credentials), used only by the PM review service.
// The PM's own sign-in token is checked first; this identity is what reads and writes the SharePoint lists,
// so project managers do not need any access to the SharePoint site.
const TENANT_ID = process.env.UPAE_TENANT_ID || "c62585ee-57f1-43a7-80b5-8efc4d57e464";
const HOST = process.env.UPAE_SP_HOST || "cacindinc.sharepoint.com";
const SITE_PATH = process.env.UPAE_SP_SITE || "/sites/CACContractManagement";

function makeGraph(opts) {
  const o = Object.assign({
    fetch: (...a) => fetch(...a),
    clientId: process.env.UPAE_GRAPH_CLIENT_ID || "",
    clientSecret: process.env.UPAE_GRAPH_CLIENT_SECRET || "",
    tenantId: TENANT_ID,
  }, opts || {});
  let tok = { at: 0, exp: 0, value: "" }, siteId = null;

  async function token() {
    if (!o.clientId || !o.clientSecret) { const e = new Error("The review service is not set up yet (the Graph client id and secret are missing)."); e.status = 503; throw e; }
    if (tok.value && Date.now() < tok.exp - 120000) return tok.value;
    const res = await o.fetch(`https://login.microsoftonline.com/${o.tenantId}/oauth2/v2.0/token`, {
      method: "POST",
      headers: { "content-type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ client_id: o.clientId, client_secret: o.clientSecret, scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" }).toString(),
    });
    const body = await res.json().catch(() => ({}));
    if (!res.ok || !body.access_token) { const e = new Error("Could not get a Graph token: " + (body.error_description || res.status).toString().slice(0, 160)); e.status = 502; throw e; }
    tok = { at: Date.now(), exp: Date.now() + (body.expires_in || 3600) * 1000, value: body.access_token };
    return tok.value;
  }
  async function call(path, init) {
    const url = path.startsWith("http") ? path : "https://graph.microsoft.com/v1.0" + path;
    const res = await o.fetch(url, Object.assign({}, init, { headers: Object.assign({ Authorization: "Bearer " + (await token()), "Content-Type": "application/json" }, (init && init.headers) || {}) }));
    if (!res.ok) { const t = await res.text().catch(() => ""); const e = new Error("Graph " + ((init && init.method) || "GET") + " failed: " + res.status + " " + t.slice(0, 160)); e.status = 502; throw e; }
    return res.status === 204 ? null : res.json();
  }
  async function site() {
    if (siteId) return siteId;
    const s = await call(`/sites/${HOST}:${SITE_PATH}`);
    siteId = s.id; return siteId;
  }
  async function listItems(list) {
    let url = `/sites/${await site()}/lists/${list}/items?expand=fields&$top=500`;
    const out = []; let guard = 0;
    while (url && guard++ < 100) { const d = await call(url); (d.value || []).forEach((x) => out.push(x)); url = d["@odata.nextLink"] || null; }
    return out;
  }
  async function getItem(list, id) { return call(`/sites/${await site()}/lists/${list}/items/${encodeURIComponent(id)}?expand=fields`); }
  async function patchFields(list, id, fields) { return call(`/sites/${await site()}/lists/${list}/items/${encodeURIComponent(id)}/fields`, { method: "PATCH", body: JSON.stringify(fields) }); }
  async function createItem(list, fields) { return call(`/sites/${await site()}/lists/${list}/items`, { method: "POST", body: JSON.stringify({ fields }) }); }
  return { listItems, getItem, patchFields, createItem };
}
module.exports = { makeGraph };
