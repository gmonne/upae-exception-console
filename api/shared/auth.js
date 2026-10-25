// Validates the Microsoft Entra ID token the browser sends (the MSAL ID token, audience = the app registration's client id).
// Uses only Node's built-in crypto: no packages to install.
const crypto = require("crypto");

const TENANT_ID = process.env.UPAE_TENANT_ID || "c62585ee-57f1-43a7-80b5-8efc4d57e464";
const CLIENT_ID = process.env.UPAE_CLIENT_ID || "b9e533a5-4da9-4999-80aa-7fc5bf2cd50a";

let jwksCache = { at: 0, keys: [] };

async function getKeys(force) {
  if (!force && Date.now() - jwksCache.at < 6 * 3600 * 1000 && jwksCache.keys.length) return jwksCache.keys;
  const res = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/discovery/v2.0/keys`);
  if (!res.ok) throw new Error("Could not fetch Microsoft signing keys: " + res.status);
  const body = await res.json();
  jwksCache = { at: Date.now(), keys: body.keys || [] };
  return jwksCache.keys;
}

function b64urlToBuf(s) { return Buffer.from(s.replace(/-/g, "+").replace(/_/g, "/"), "base64"); }

// Returns { email, name, oid } or throws an Error with .status = 401
async function verifyToken(authHeader, opts) {
  const o = Object.assign({ tenantId: TENANT_ID, clientId: CLIENT_ID, getKeys }, opts || {});
  const fail = (m) => { const e = new Error(m); e.status = 401; return e; };
  const m = /^Bearer\s+(.+)$/i.exec(authHeader || "");
  if (!m) throw fail("Missing sign-in token.");
  const parts = m[1].split(".");
  if (parts.length !== 3) throw fail("Malformed token.");
  let header, payload;
  try {
    header = JSON.parse(b64urlToBuf(parts[0]).toString("utf8"));
    payload = JSON.parse(b64urlToBuf(parts[1]).toString("utf8"));
  } catch (e) { throw fail("Malformed token."); }
  if (header.alg !== "RS256") throw fail("Unsupported token algorithm.");

  let keys = await o.getKeys(false);
  let jwk = keys.find((k) => k.kid === header.kid);
  if (!jwk) { keys = await o.getKeys(true); jwk = keys.find((k) => k.kid === header.kid); }
  if (!jwk) throw fail("Unknown signing key.");
  const key = crypto.createPublicKey({ key: jwk, format: "jwk" });
  const ok = crypto.verify("RSA-SHA256", Buffer.from(parts[0] + "." + parts[1]), key, b64urlToBuf(parts[2]));
  if (!ok) throw fail("Invalid token signature.");

  const now = Math.floor(Date.now() / 1000);
  if (payload.exp && now > payload.exp + 60) throw fail("Token expired. Please sign in again.");
  if (payload.nbf && now < payload.nbf - 60) throw fail("Token not yet valid.");
  if (payload.aud !== o.clientId) throw fail("Token is for a different application.");
  if (payload.tid !== o.tenantId) throw fail("Token is from a different organization.");
  if (!String(payload.iss || "").startsWith("https://login.microsoftonline.com/" + o.tenantId)) throw fail("Unexpected token issuer.");
  return { email: String(payload.preferred_username || payload.email || payload.upn || "").toLowerCase(), name: payload.name || "", oid: payload.oid || "" };
}

module.exports = { verifyToken };
