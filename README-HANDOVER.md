# Swimlane Studio — Handover Documentation (First Pass)

Prepared by William Sands · Draft for review · 2026-08-20

This is a first-pass answer to the handover documentation request, covering
the ten areas requested. Facts marked **CONFIRMED** are verified directly
against the live repository, config files, or code comments as of this
date. Facts marked **NEEDS INPUT** are genuinely outside what's visible
from the codebase (organisational, licensing, or personnel questions) and
need William's or another stakeholder's input to complete. Nothing below
is guessed and presented as fact — where something isn't known, it says so.

---

## 1. Infrastructure & Hosting

**Where it's hosted.** GitHub Pages, serving a static build of the app.
Repository: `https://github.com/sandmansss45/swimlane`. Live URL:
`https://sandmansss45.github.io/swimlane/`. **CONFIRMED.**

**Hosting architecture and dependencies.** This is a pure client-side
single-page app (React) with no backend or server component of its own.
It runs entirely in the visiting browser. All dynamic data comes directly
from Microsoft Graph API calls made *from the browser*, authenticated via
MSAL.js against Entra ID. There is no database — SharePoint Online lists
are the only data store. Runtime dependencies: GitHub Pages (hosting),
Microsoft Entra ID (auth), Microsoft Graph / SharePoint Online (data).
Build-time only: the npm registry (dependency resolution during `npm
install`/CI, not needed at runtime). **CONFIRMED.**

**Development, test, and production environments.** Two exist:
- **Local dev** — `npm run dev` (Vite dev server, `http://localhost:5173`),
  against either mock fixture data or the real SharePoint site.
- **Production** — GitHub Pages, auto-deployed from `main`.

There is **no separate test/staging environment** and no formal
environment promotion process. A push to `main` goes straight to
production. **CONFIRMED — this is a real gap, see Risks.**

**Deployment process.** Fully automated via GitHub Actions
(`.github/workflows/deploy.yml`): a push to `main` triggers `npm ci &&
npm run build`, the built `dist/` folder is uploaded as a Pages artifact,
and `actions/deploy-pages` publishes it. No manual steps, no approval
gate. Typically live within 1–2 minutes of the push. **CONFIRMED.**

**Release management process.** None formal exists — no version
tagging, no changelog, no staged rollout, no rollback mechanism beyond
`git revert` (which itself triggers a new auto-deploy). **CONFIRMED gap.**

**Monitoring and alerting.** None. No Application Insights, no error
tracking (e.g. Sentry), no uptime monitoring, no alerting of any kind.
The only way to know a deploy failed is to check the repo's **Actions**
tab by hand. **CONFIRMED gap.**

**Logging capabilities.** Browser console only — a handful of
`console.log`/`console.error` calls in the data-loading code (e.g.
confirming a fetched row count) for manual diagnosis. Nothing is
captured or aggregated anywhere; it only exists in whoever's own
DevTools at the time. **CONFIRMED.**

**Licences, subscriptions, and external dependencies.** All npm
dependencies are open-source (React, Vite, Fluent UI, MSAL.js,
html-to-image, jsPDF — see `package.json`). No paid third-party API or
service is used anywhere in the app — this was a deliberate, explicit
constraint (a proposed AI feature was scoped but not built specifically
to keep the app free to run). **CONFIRMED — no paid dependencies in the
app itself.** GitHub account/organisation plan tier and Microsoft
365/Entra ID tenant licensing are **NEEDS INPUT** — not visible from the
repository.

**Expected performance and limitations.** Client-side rendered; load
time scales with how many rows exist in the underlying SharePoint lists
(Graph's own server-side pagination is handled, but the app loads full
list contents, not a paged view). The production JS bundle is
approximately 1.3MB minified, in a single chunk (Vite's own build output
flags this as larger than its recommended threshold — no code-splitting
has been done). PDF export renders the entire visible diagram client-side
and can be slow/memory-heavy on a very wide flow. No formal performance
SLA exists. **CONFIRMED (bundle size from build output; no SLA exists).**

**What happens if a key component becomes unavailable:**
- **GitHub Pages down** → the app is entirely unreachable. No fallback
  hosting exists.
- **Entra ID / Microsoft Graph unavailable or misconfigured** → real
  sign-in fails, but the "Use mock data (no sign-in)" option on the
  sign-in screen still works, showing the tool's behaviour against
  fixture data (not real data) so it can still be demonstrated.
- **A specific SharePoint list becomes unavailable, renamed, or
  permission is revoked** → most secondary lists (Process Group labels,
  Process ID labels, Process ID locks, Swimlane comments) fail open —
  the rest of the app keeps working, just without that one feature. The
  core process-step list (**Master File**) failing is a harder failure —
  the main Process Flows view can't load without it.

**CONFIRMED**, by design (`Promise.allSettled` used deliberately in the
main data-loading path for this reason).

---

## 2. Authentication & Security

**Entra ID configuration.**
- Tenant: Quantum Leap Energy, tenant ID `f75d6062-398e-493a-9af8-5b532b88d594`.
- App registration: **"Swimlane Studio"**, client (application) ID
  `b654eeeb-7451-49cf-94e4-d089592de6e5`.
- Registered redirect URIs: `http://localhost:5173/` (dev) and
  `https://sandmansss45.github.io/swimlane/` (production). **If the
  repo's owner or name ever changes, the production URL changes too, and
  this redirect URI must be updated in Entra or sign-in breaks
  (AADSTS50011).**

**CONFIRMED** (from `src/swimlane/auth/authConfig.ts` and `AUTH_SETUP.md`).

**Updating the redirect URI after a GitHub transfer (step-by-step).**
Nothing else about this Entra app registration needs to change — same
tenant, same Client ID, same permissions, all already owned by QLE. This
is the only step required:
1. Go to `portal.azure.com` (or `entra.microsoft.com`) and sign in with
   an account that has admin rights on the QLE tenant.
2. Search for **App registrations** in the top search bar.
3. Click **Swimlane Studio** in the list.
4. Click **Authentication** in the left-hand menu.
5. Under the **Web** platform's redirect URIs, click **Add URI** and
   enter the new address exactly: `https://<new-owner>.github.io/swimlane/`
   (replace `<new-owner>` with whatever GitHub organisation the repo was
   transferred to — see Section A below).
6. Click **Save** at the top of the page.
7. Once someone's confirmed sign-in works at the new URL, come back and
   remove the old `https://sandmansss45.github.io/swimlane/` entry the
   same way (select it, click the bin icon, **Save**) so a personal
   account URL doesn't stay valid indefinitely.

**API permissions and consent.** Microsoft Graph, delegated: `User.Read`
and `Sites.Selected`. Deliberately **not** `Sites.ReadWrite.All` — this
app can only ever read/write the one SharePoint site it's been
explicitly granted access to, never anything else in the tenant. Admin
consent has been granted. `Sites.Selected` additionally required a
one-time Graph API call (separate from the app registration UI) granting
this specific app write access to the specific "Swimlane Studio" site —
documented step-by-step in `AUTH_SETUP.md`. **CONFIRMED.**

**Service accounts.** None. Every user signs in with their own Entra
identity; there is no shared or service account anywhere in the app.
**CONFIRMED.**

**Conditional Access considerations.** **NEEDS INPUT** — not configured
at the app level and not visible from the codebase; whatever Conditional
Access policies apply tenant-wide apply as normal.

**Administrative access requirements.** Changing the Entra app
registration (redirect URIs, permissions) requires Application
Administrator (or Global Administrator). Granting `Sites.Selected` site
access requires SharePoint Administrator or Global Administrator.
**CONFIRMED** (standard Entra/SharePoint role requirements for these
actions).

**Security considerations and known constraints.**
- This is a **client-side-only app with no server enforcing
  permissions**. Any user who can sign in and who has access to the
  SharePoint site can read/write anything the app's Graph permissions
  allow.
- The swimlane **lock** feature is a social/audit control, not a
  technical permission barrier, by explicit design decision: any signed-in
  user can lock or unlock a swimlane. Accountability comes from every
  action being attributed and permanently logged, not from being
  technically blocked.
- There is currently **no role-based access control** — no "view only"
  vs "can edit" distinction exists anywhere in the app. Every signed-in
  user with site access can edit everything that isn't currently locked.
  This was explicitly raised as a real requirement by a stakeholder and
  scoped in discussion, but **not implemented** — a genuine, current gap
  (see Risks and Roadmap).

**CONFIRMED** (all by direct code inspection and explicit design
decisions made during development).

**Access management process.** Who *can* sign in is governed entirely by
standard tenant and SharePoint site permissions — nothing inside the app
itself manages user access. **CONFIRMED.**

---

## 3. Data Architecture

**SharePoint sites used.** One dedicated site:
`qleapenergy.sharepoint.com/sites/SwimlaneStudio` ("Swimlane Studio"),
created specifically for this app. **CONFIRMED.**

**Lists, libraries, and data stores, and their purpose:**

| List | Site | Purpose | Status |
|---|---|---|---|
| **Master File** (renamed from "9.6 tester" 2026-08-19) | Swimlane Studio | The core process-step data every swimlane is drawn from | CONFIRMED, real data |
| **Process Group labels** | Swimlane Studio | User-created names for Process Groups not already in the app's built-in APQC reference table | CONFIRMED, created |
| **Process ID labels** | Swimlane Studio | Same idea, one level down — names for empty Process ID "shells" before real steps exist | CONFIRMED, created |
| **Process ID locks** | Swimlane Studio | Append-only audit trail of every swimlane lock/unlock | CONFIRMED, created |
| **Swimlane comments** | Swimlane Studio | Append-only feedback log (the "Improvements" feature) | CONFIRMED, created |
| **QLE Existing Organisation** | *(pre-existing, not owned by this app)* | Read-only source for employee job titles/departments | CONFIRMED, read-only |
| **risk register data** | *(pre-existing, not owned by this app)* | The enterprise Risk Register — read-only, linked into steps | CONFIRMED, read-only |

**Master File columns:** APQC Title, Process Description, Process Step
ID, Process Step Name, Action Type, Action, Action Description,
ResponsibleJobTitle, Region, ShapeOverride, ManualOrder, DependsOn, Edge
Labels, Linked Risks, SOP Link, Delegation of Authority Link — all
**CONFIRMED** present on the real list.
`RiskLevelOverride` — a retired manual risk-level flag, superseded by
Linked Risks (a real Risk Register link, the app's one actual risk
indicator) — has been fully removed from the app, and the column has
been deleted from the real list. **CONFIRMED, done.**

**Data ownership and maintenance.** Master File, Process Group labels,
Process ID labels, Process ID locks, and Swimlane comments are all
owned by this app (created specifically for it, on its own site). QLE
Existing Organisation and risk register data are owned elsewhere —
standing enterprise lists this app only ever reads from, never writes
to. **CONFIRMED.**

**Metadata structures.** Per-step "created by / last modified by / when"
is read directly from SharePoint's own native, system-managed item
metadata rather than a duplicate custom column — SharePoint already
tracks this accurately for free. **CONFIRMED**, deliberate design choice.

**Relationships between data sets.** `DependsOn` encodes step-to-step
sequence, by the step's **original row number in the source data** (not
by any ID — header counted as row 1). `Linked Risks` (once the column
exists) will encode step-to-risk links by the real Risk Register's
SharePoint item ID plus a manually chosen severity. `Region` splits one
Process ID into several parallel swimlanes (UK/US/SA/Global).
**CONFIRMED.**

**Source systems and integrations.** Only Microsoft Graph →
SharePoint Online. No other system integration exists — process steps
that *describe* activity in other systems (e.g. NetSuite) are just text
descriptions of a real-world process; there is no live NetSuite
integration. **CONFIRMED.**

**How data synchronisation works.** There is no batch or scheduled sync.
The app reads live from Graph on every page load and writes directly
back to SharePoint via Graph the instant an edit is saved. No caching
layer, no offline mode. **CONFIRMED.**

**Recovery process for data inconsistency or sync failure.** No custom
process exists. SharePoint Online's own built-in per-item version
history is the only safety net currently available — nothing bespoke has
been built. **CONFIRMED gap** — see Recovery & Business Continuity.

---

## 4. Solution Architecture

**High-level architecture.** Browser (React SPA) → MSAL.js → Entra ID
(auth) → Microsoft Graph API → SharePoint Online (7 lists across 2
sites). No backend/API layer of its own — Graph *is* the API layer.

**Data model.** Central entity is `IProcessStep` (one row = one action
in a process). Supporting entities: `IEmployee`, `IRiskStatement`,
`IProcessGroupLabel`, `IProcessIdLabel`, `IProcessIdLock`,
`ISwimlaneComment`. Full definitions in `src/swimlane/models/`.

**Integration map.** Browser → MSAL.js (Entra ID, OAuth2/OIDC) →
Microsoft Graph API (`https://graph.microsoft.com/v1.0`) → SharePoint
Online lists (see Data Architecture table above). That's the entire
integration surface — nothing else.

**Application structure.** Single-page app with five top-level views:
**Process Flows** (the swimlane diagrams — drill down Category → Process
Group → Process ID → swimlane), **Employees** (directory, read-only),
**Risk Register** (read-only), **Audit** (per-step change history plus
lock history), **Improvements** (cross-swimlane comment feed).

**Source code structure.**
```
src/main.tsx                        entry point
src/App.tsx                         auth gate / mock-mode switch
src/swimlane/
  components/
    SwimlaneStudio.tsx              main app shell + state
    SwimlaneCanvas.tsx              the diagram grid renderer itself
    ...                             supporting UI components
  services/
    IDataService.ts                 the interface both data sources implement
    MockDataService.ts              in-memory fixture data
    GraphDataService.ts             real Microsoft Graph / SharePoint
  models/                           TypeScript interfaces (the data model)
  utils/                            pure logic - dependency resolution,
                                     arrow-routing math, APQC hierarchy,
                                     CSV import
  auth/                             MSAL config, Graph client
```

**Key design decisions and rationale.**
- **The diagram is a structured lane × column grid, not a free-form
  canvas.** Every step gets its own column; lanes are job titles. This
  was a deliberate choice over a full diagramming-tool rebuild — it's
  what makes automatic arrow routing, dependency-order validation,
  region-splitting, and CSV round-tripping all work without a
  from-scratch rewrite.
- **DependsOn tokens are row numbers, matching the real source data's
  own convention** — not redesigned to use IDs, so a real export
  continues to import correctly.
- **Mock/demo data is real (anonymisation aside) accounts-payable
  process data**, not synthetic placeholder data, by explicit request —
  so demoing in mock mode shows the actual real flow, not a toy example.
- **`Sites.Selected` over `Sites.ReadWrite.All`** as a deliberate,
  narrower security boundary — this app can never touch anything in the
  tenant beyond its own one site.

**Technology stack.** React 19, TypeScript, Vite 8, Fluent UI v8
(`@fluentui/react`), MSAL.js (`@azure/msal-browser`,
`@azure/msal-react`), `html-to-image` + `jsPDF` (PDF export), Sass (CSS
Modules). Build/lint: `tsc`, Vite, `oxlint`. No test framework is
installed. Full list in `package.json`.

---

## 5. Operations & Support

**Support process.** None formally established — no ticketing system,
no defined support tier, no SLA. William has been the de facto sole
support contact. **This gap is exactly what this handover is meant to
close.**

**Common issues and resolutions:**
- *Sign-in fails with AADSTS50011* → the redirect URI Entra has
  registered no longer matches the live URL (usually after a repo
  rename/transfer) → update the redirect URI in the Entra app
  registration.
- *A feature works in mock mode but silently doesn't save in real mode*
  → almost always means a SharePoint column the code expects doesn't
  exist yet on the real list → check the `TODO-CONFIRM` comments in
  `GraphDataService.ts`, then create the column with the exact name/type
  documented there.
- *"No list titled X found on the site" error* → a real SharePoint list
  was renamed without updating the matching constant in
  `GraphDataService.ts` → update that constant to the list's current
  real display name.

**Troubleshooting approach.** Browser DevTools console is the primary —
effectively only — diagnostic tool. Several data-loading calls log the
row count they fetched, useful for confirming a Graph call actually
reached the list it was meant to.

**Known issues.** None outstanding on the Master File schema — Linked
Risks, SOP Link, and Delegation of Authority Link are all confirmed
present, and the retired `RiskLevelOverride` column has been removed
from both the app and the real list.

**Known limitations.** No automated tests. No monitoring/alerting. No
formal release process. No role-based access control. CSV import
doesn't detect duplicate IDs (a deliberate scope decision — a human is
expected to review the result). Undo covers only the single most recent
action, not a full history. PDF export can be slow on a very wide flow.

**Escalation paths.** **NEEDS INPUT / not yet defined** — part of what
the planned governance model should establish.

**Operational activities requiring regular attention.** Periodically
checking the repo's Actions tab after a push to confirm the deploy
actually succeeded — nothing alerts automatically if one fails. No other
recurring operational task currently exists.

**Maintenance requirements.** npm dependency updates aren't on any
schedule. The Entra app registration's redirect URIs need updating any
time the app's hosting URL changes.

---

## 6. Recovery & Business Continuity

**Backup arrangements.** None custom-built by this app. Relies entirely
on SharePoint Online's own built-in versioning/retention.
**NEEDS INPUT** on what the tenant's actual retention/backup policy is —
that's a Microsoft 365 administration question, not visible from the app.

**Recovery procedures.** None formally defined for this app specifically.

**Disaster recovery considerations.** If GitHub itself were unavailable,
there is no alternate hosting standing by. The source code and its full
history are safely held in the git repository, though (clonable by
anyone with access) — re-deploying elsewhere would mean cloning, pushing
to a new host, and updating the Entra redirect URI to match.

**Recovery dependencies.** The git repository (source of truth for the
app), SharePoint Online (source of truth for all data), and a correctly
configured Entra app registration (redirect URI matching wherever the
app is actually being served).

**Recovery testing undertaken.** None to date.

**Expected recovery approach following a major failure.** Re-deploy the
app (clone → push to a GitHub repo with Pages enabled → update the Entra
redirect URI). Data loss would only occur if SharePoint Online itself
lost data, which falls under Microsoft's own service responsibility, not
this app's.

---

## 7. Access & Ownership

**Current administrators, current owners, and who owns what** (repository,
application, SharePoint site, source code) — **all NEEDS INPUT.** I can
confirm the *technical* configuration (who the app registration trusts,
what site it's scoped to), but not the *organisational* answer of who
currently holds admin rights to the GitHub repository, the Entra app
registration, or the SharePoint site. That needs William's direct input,
ideally captured as an explicit RACI as part of the governance model.

**One fact worth flagging directly here rather than leaving implicit:**
the GitHub repository (`sandmansss45/swimlane`) appears to sit under a
**personal GitHub account**, not an organisation-owned one. This is a
real, near-term risk — see Risks below.

---

## 8. Product Roadmap

**Current status.** Functionally complete for its core purpose —
visualising and editing swimlane process flows sourced from SharePoint —
and was under active, frequent development right up to this handover
(84 commits in the repository's history).

**Technical debt.** No automated test suite. No code-splitting (single
~1.3MB JS bundle). A couple of SharePoint column names are still
unconfirmed guesses rather than verified facts. No formal
release/versioning process.

**Planned enhancements (raised, discussed, explicitly not yet built):**
- **Role-based access control** (view-only vs. can-edit) — raised
  directly by a stakeholder, scoped in discussion, not implemented.
- **"Published" vs. "In-progress" staging**, with function-lead-only
  edit rights even on published items — raised as an idea, explicitly
  deferred, never scoped in detail. Closely related to both the lock
  feature and the access-control ask above.
- **Real AI integration** — discussed (in-browser free-tier options vs.
  a small hosted API), explicitly **not** built, and explicitly
  constrained to free/no-added-cost options at the user's request. No
  direction has been committed to.

**Outstanding work.** None on the schema front — see Known Issues above.

**Future architecture considerations.** If role-based access control is
built, it needs to go beyond an in-app UI toggle — a restriction enforced
only in the browser isn't real security, since the same Graph API access
would still need locking down at the SharePoint permission level too.

**AI roadmap and current thinking.** No committed roadmap. Two options
were scoped and discussed (an in-browser, fully client-side model with
no per-use cost; or a minimal serverless API with its own hosting cost)
but neither has been chosen or built. This is a live open decision for
whoever inherits product direction.

---

## 9. Single Points of Failure

Everything below currently depends on William's own knowledge or direct
involvement:

- **Deployment** — pushing to `main` is the only way a change goes live;
  no one else has done this yet.
- **Source code knowledge** — William is currently the only person who
  has worked in this codebase. (This document, plus a new `CLAUDE.md` in
  the repo root written specifically to orient a new contributor quickly,
  are meant to close this gap.)
- **App registration management** — Entra changes require both
  Application Administrator rights *and* knowing exactly what's already
  configured (documented in `AUTH_SETUP.md` and this document, but not
  yet exercised by anyone else).
- **PowerShell scripts** — none exist as persisted, reusable scripts.
  Verification during development used one-off Playwright scripts
  written to a temporary folder and deleted after use — nothing was kept
  as a reusable tool.
- **Administrative tasks** — granting `Sites.Selected` site access (the
  one-time Graph permission grant) requires a SharePoint/Global Admin
  following the steps in `AUTH_SETUP.md`; only done once so far.
- **Manual processes / workarounds** — the portable Node.js/Git installs
  used on William's own development machine (no admin rights to install
  system-wide) are a personal workaround, not something the app depends
  on — a new contributor should do a normal system-wide install (see the
  VS Code setup guide below).
- **Future design intentions** — captured in the Roadmap section above;
  nothing exists only in conversation that isn't now written down here
  or in the code's own comments (which are unusually thorough throughout
  this codebase, specifically so context doesn't depend on asking
  someone).
- **Configuration not currently documented** — GitHub/Microsoft 365
  licensing tiers, Conditional Access policies, and who else (if anyone)
  currently holds admin access to the repo, app registration, or
  SharePoint site.

---

## 10. Risks, Assumptions & Recommendations

**Key risks.**
1. The GitHub repository is under a personal account, not an
   organisation-owned one — a single point of failure for both the
   source code and the deploy pipeline.
2. No monitoring or alerting — a broken deploy or a broken Graph
   integration could go unnoticed until a user reports it.
3. No role-based access control — every signed-in user can edit every
   swimlane that isn't currently locked; locking is a social/audit
   control, not a technical barrier.
4. No automated tests — every future change relies entirely on manual
   verification.
5. A couple of SharePoint column names are unconfirmed guesses rather
   than verified facts (risk of a feature silently not persisting in
   real mode until specifically checked).

**Assumptions built into the solution.**
- Every user who can sign in via the tenant and has site access is
  trusted to edit — there is no finer-grained trust model.
- SharePoint Online is always treated as the live, authoritative data
  source — the app never caches and has no offline mode.
- Scoping the app's Graph permission to one SharePoint site
  (`Sites.Selected`) was judged an adequate security boundary given the
  app's current risk profile and lack of in-app permission tiers.

**Areas of concern.** Same as Key risks, above.

**Governance considerations.** The planned governance model should
explicitly cover: role-based access, change management for the
self-service "add a new APQC category/group" flows, and what "locking a
swimlane" is actually meant to certify (today it's a social convention —
anyone can do it — not an enforced review/sign-off step).

**Recommended next steps.**
1. Transfer the GitHub repository to an organisation-owned account.
2. Stand up basic monitoring — at minimum, a notification on GitHub
   Actions deploy failures.
3. Decide and scope role-based access control before it's needed for
   real.
4. Explicitly document who holds admin rights to each of: the GitHub
   repository, the Entra app registration, and the SharePoint site.

**Anything a new owner should know that may not be obvious.** This
codebase's inline comments are unusually thorough and consistently
explain *why* a decision was made, not just what the code does — reading
them is often faster and more reliable than asking someone. `CLAUDE.md`
(added to the repo root as part of this handover) is written specifically
to orient a new contributor — human or AI-assisted — quickly, and to
keep that same standard going forward.

---

## A. Transferring the GitHub repo / GitHub Pages hosting to someone else

The app is *already* on GitHub with Pages already live, so "handing it
off" is really one of these three, not a from-scratch setup:

**Step 0 — If there's no company-owned GitHub organisation yet.** The
whole point of this transfer is getting off a personal account, so
transferring into someone else's personal account instead doesn't
actually fix that — it needs to land in a proper GitHub Organisation
that IT/engineering leadership controls, not one individual.
1. Go to `github.com`, signed in as whoever will administer it.
2. Click the **+** icon (top-right) → **New organization**.
3. Pick the **Free** plan unless there's a specific reason for a paid
   tier — it's enough for what this app needs.
4. Enter an organisation name (e.g. something identifying the company,
   like `quantum-leap-energy`) and a contact email.
5. Follow the remaining prompts (email verification, optionally
   inviting other admins now or later).
6. The organisation now exists and can be the destination for Option 1
   below.

**Option 1 — Transfer the existing repository (recommended).** Keeps the
full commit history, the Pages configuration, and the Actions run
history intact.
1. On the repo page → **Settings** → scroll to **Danger Zone** → **Transfer
   ownership**.
2. Enter the new owner's GitHub username or organisation name.
3. The new owner accepts the transfer invitation.
4. **The live URL changes** to `https://<new-owner>.github.io/swimlane/`
   (or a different path if they also rename the repo). **This means the
   Entra app registration's redirect URI must be updated to match** —
   see Section 2 above — or sign-in will break immediately.

**Option 2 — Add someone as a collaborator**, if joint or interim access
is wanted rather than a full handover: **Settings** → **Collaborators** →
**Add people**. No URL change, no Entra update needed.

**Option 3 — A completely fresh copy under a different account** (only
if the intent is a genuinely separate copy, e.g. a personal fork or
demo, not the real handover): this is exactly what `DEPLOY.md` in this
repo already documents step-by-step — create a new empty repo, `git
remote add origin <new-url>`, `git push`, enable Pages under **Settings
→ Pages → Source: GitHub Actions**. Same redirect-URI caveat applies.

**Recommendation:** given the personal-account risk flagged in Section
10, Option 1 into an organisation-owned GitHub account is the right
long-term move, done once, rather than repeatedly copying the repo
around.

---

## B. Making changes going forward — and why Claude Code specifically

This app was built collaboratively with Claude Code across many focused
sessions, not written and then documented after the fact. That shows up
directly in the codebase: comments throughout consistently explain *why*
a decision was made (often citing a real bug it fixed, or an explicit
user decision with a date) rather than just what the code does — exactly
the kind of context an AI coding assistant (or a new human developer)
needs to make a safe change without re-deriving it from scratch by
reading every file. `CLAUDE.md` in the repo root exists specifically to
hand that context over automatically.

**Recommended workflow for any future change:**
1. Open the repo in VS Code with Claude Code installed (steps below).
2. Describe the change in plain language. Claude Code reads `CLAUDE.md`
   automatically for project/architecture context, and can be pointed at
   this document for the operational picture.
3. Insist on `npm run build` succeeding with zero errors, and on the
   actual feature being verified live — ideally in a browser, in mock
   mode where no credentials are needed — not just "the code compiles."
4. Review the diff before committing. A push to `main` deploys to
   production immediately with no approval gate (see Section 1) — there
   is no safety net between "committed" and "live."

Changes are still possible without Claude Code for anyone comfortable
reading React/TypeScript directly, but a lot of the app's non-obvious
behaviour — why arrows route the way they do, why `DependsOn` uses row
numbers instead of IDs, why swimlane locking is a social control and not
a technical one — lives in comments that are easy to miss without
something deliberately reading the whole file first.

### Setting up Claude Code in VS Code

1. Install **VS Code**: https://code.visualstudio.com
2. Install **Node.js** (LTS): https://nodejs.org — a normal system-wide
   install, not the portable workaround used on William's own machine
   during development (that workaround exists only because that specific
   machine had no admin rights).
3. Install **Git for Windows** (or the equivalent for macOS/Linux):
   https://git-scm.com — again, a normal system-wide install.
4. Clone the repository:
   ```
   git clone https://github.com/sandmansss45/swimlane.git
   cd swimlane
   npm install
   npm run dev
   ```
   Confirm it opens at `http://localhost:5173` and that "Use mock data
   (no sign-in)" loads the app without needing any credentials.
5. Install **Claude Code**: https://claude.com/claude-code — either the
   VS Code extension from the Marketplace, or the CLI per Anthropic's own
   setup instructions. Requires signing in with a Claude account that has
   Claude Code access (a paid Claude plan or API access) — **this is a
   genuine external cost decision for whoever takes this on**, separate
   from the app's own "must stay free to run" constraint, which is about
   the *app itself* never calling a paid API, not about the tooling used
   to develop it.
6. Open the cloned `swimlane` folder in VS Code. Claude Code
   automatically reads `CLAUDE.md` in the repo root for project context —
   it doesn't need to be told what the project is or how it's built.
7. Read this document (`README-HANDOVER.md`) before making any change
   that touches authentication, hosting, or the SharePoint data model —
   it's the map of what's confirmed vs. what's still a guess.

---

*This is a first pass, written to be worked through together rather than
treated as final — flagged gaps above are exactly the kind of thing this
document is meant to surface for that conversation.*
