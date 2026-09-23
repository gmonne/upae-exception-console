# UPAE Exception Console

A concept UI for the **Unified Payroll Accuracy Engine (UPAE)** — an AI-enhanced payroll validation platform that proactively identifies payroll discrepancies (missing hours, incorrect union classifications, overtime miscalculations, job allocation errors, duplicate entries, and union work-rule violations) before payroll is processed.

This repo holds the front-end mock-up used for team walkthroughs and pilot planning, plus the SharePoint/Microsoft Graph integration layer that backs its Rule Library and Admin screens with a real, shared data store.

**Live demo:** https://gentle-field-0b00f5c0f.6.azurestaticapps.net *(auto-deployed from `main` — see [Deployment](#deployment))*

## What's in this console

| Screen | What it shows |
|---|---|
| **Dashboard** | Sample exception queue and validation run, illustrating what a payroll reviewer sees day to day |
| **Insights** | Illustrative exception-detection analytics (volume, $ impact, detection lead time), plus a live **Rule library & reference data coverage** panel showing real counts and status breakdowns computed from the current Unions / Labor Types / Rules data |
| **Rule library** | Browse, add, and edit structured union pay rules (regular time, OT 1.5x, OT 2x, Night Differential) against the [Rule Definition Schema](#rule-definition-schema), with per-tier "Not applicable to this union" switches and worked examples |
| **Admin: Unions & Labor Types** | Manage the reference lists that labor type codes map to |

Screens marked "Merged" or "Roadmap" in the nav are folded into Dashboard or intentionally out of scope for this concept.

## Live data via SharePoint

The Rule Library and Admin screens can run two ways:

- **Demo mode** — no sign-in, uses in-memory sample data seeded in `index.html`; nothing persists across a page refresh.
- **Live mode** — sign in with a Microsoft 365 account (a sign-in gate appears on load) to load and save against five SharePoint lists (`Unions`, `LaborTypes`, `Rules`, `RuleLaborTypes`, `RuleExamples`) via Microsoft Graph.

See [`sharepoint-list-schema.md`](sharepoint-list-schema.md) for the exact list/column schema and [`sharepoint-integration-wiring-guide.md`](sharepoint-integration-wiring-guide.md) for what's wired up vs. still on the roadmap (currently: Rule save is fully live; Admin add/delete and the Add Example modal are still demo-only).

### Auth setup

Sign-in uses [MSAL.js](https://github.com/AzureAD/microsoft-authentication-library-for-js) against an Azure AD (Microsoft Entra ID) app registration:

- Platform type must be **Single-page application (SPA)**, not "Web"
- Redirect URI must match wherever this is hosted (the Static Web Apps URL above, or `http://localhost:<port>` for local testing)
- API permission: `Sites.ReadWrite.All` (delegated), with admin consent granted
- Signed-in users need at least **Contribute** access on the target SharePoint site

Configuration (tenant ID, client ID, site path, list names) lives at the top of [`sharepoint-integration.js`](sharepoint-integration.js).

The MSAL library itself is vendored locally as `msal-browser.min.js` rather than loaded from a CDN, since `alcdn.msauth.net` isn't reachable from every network this has been tested on and a failed CDN load silently breaks the sign-in button.

## Running locally

This is a static site — no build step, no dependencies to install.

```bash
git clone <this-repo>
cd upae-exception-console
python3 -m http.server 8080
# open http://localhost:8080
```

If you want to test live SharePoint sign-in locally, add `http://localhost:8080` as a redirect URI on the app registration first.

## Deployment

Deploys automatically via GitHub Actions ([`.github/workflows/azure-static-web-apps-gentle-field-0b00f5c0f.yml`](.github/workflows/azure-static-web-apps-gentle-field-0b00f5c0f.yml)) to Azure Static Web Apps on every push to `main`. No build step — the whole repo root is uploaded as-is (`app_location: "/"`).

## Repo structure

```
index.html                                  Single-page mock-up (all markup/CSS/JS except the SharePoint layer)
sharepoint-integration.js                   Auth + Graph API reads/writes, sign-in gate wiring
msal-browser.min.js                         Vendored MSAL.js library (self-hosted, see Auth setup)
favicon.ico, favicon-*.png                  App icon
sharepoint-list-schema.md                   SharePoint list/column definitions
sharepoint-integration-wiring-guide.md      What's wired to SharePoint vs. still demo-only
rule-definition-schema.md                   Field-by-field format for a union pay rule
```

## Status

Concept/pilot stage — sample data throughout is illustrative, not live payroll figures. See the project's other docs for the current build order and open questions (labor type → union crosswalk gaps, Night Differential precedence rules for additional unions, etc.).
