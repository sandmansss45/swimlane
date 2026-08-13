# Entra app registration + site permission setup

## Status (as of 2026-08-13)

- App registration exists: **Swimlane Studio**, client ID `b654eeeb-7451-49cf-94e4-d089592de6e5`.
- Site created: **qleapenergy.sharepoint.com/sites/SwimlaneStudio** - confirmed
  and filled into `authConfig.ts`.
- Steps 1 and 2 below (permission set to `Sites.Selected` + site-specific
  access grant) - reported done by Method Group, not yet verified end to end
  with a real sign-in.
- **Still outstanding, separate from auth**: `PROCESS_LIST_TITLE` and
  `EMPLOYEES_LIST_TITLE` in `src/swimlane/services/GraphDataService.ts` are
  still literal placeholder strings. Even with sign-in working, the app will
  fail to load real data until the actual SharePoint **Lists** (not the
  default Documents library) exist in this site and their real titles are
  filled in there.

## 1. Confirm the app registration's permission is `Sites.Selected`

If it's still showing `Sites.ReadWrite.All` from the earlier request:

1. Entra admin center (entra.microsoft.com) → **App registrations** → **Swimlane Studio** → **API permissions**.
2. Remove `Sites.ReadWrite.All` if present. Add **`Sites.Selected`** (Microsoft Graph, delegated). `User.Read` stays as-is.
3. Click **Grant admin consent for Quantum Leap Energy** → **Yes**.

Note: admin consent here only allows the *app* to request `Sites.Selected`
tokens — it does **not** by itself give the app access to any specific site.
That's step 2.

## 2. Grant the app access to the new "Swimlane Studio" site

This is the extra step `Sites.Selected` needs that `Sites.ReadWrite.All`
wouldn't have — a one-time Graph API call scoping this specific app to this
specific site. Needs a SharePoint Administrator or Global Administrator.
Easiest way (no PowerShell install needed):

1. Go to https://developer.microsoft.com/graph/graph-explorer and sign in
   with an admin account.
2. If prompted, consent to the `Sites.FullControl.All` permission for Graph
   Explorer itself (only needed to run this one call, not part of our app).
3. Find the new site's ID: run a GET request to
   `https://graph.microsoft.com/v1.0/sites/qleapenergy.sharepoint.com:/sites/<new-site-url-name>`
   and copy the `id` field from the response.
4. Run a POST request to `https://graph.microsoft.com/v1.0/sites/<site-id>/permissions`
   with this body:
   ```json
   {
     "roles": ["write"],
     "grantedToIdentities": [
       {
         "application": {
           "id": "b654eeeb-7451-49cf-94e4-d089592de6e5",
           "displayName": "Swimlane Studio"
         }
       }
     ]
   }
   ```
5. A 201 response confirms it — the app can now read/write that one site and
   nothing else.

(Alternative if your team prefers PowerShell: PnP PowerShell's
`Grant-PnPAzureADAppSitePermission -AppId b654eeeb-7451-49cf-94e4-d089592de6e5 -Site <site-url> -Permissions Write`
does the same thing.)

## 3. Tell me the new site's URL

Once the site exists, send me its URL — I'll fill in
`SHAREPOINT_SITE_HOSTNAME`/`SHAREPOINT_SITE_PATH` in `authConfig.ts` (currently
a placeholder) and we can test real sign-in end to end.

## 4. Redirect URI - CORRECTION, real sign-in tested 2026-08-13

Testing the actual sign-in flow against the live GitHub Pages site turned up
a real bug: it landed on `sandmansss45.github.io/` (bare domain, GitHub's
404 page) instead of the app. Cause: GitHub Pages serves this project repo
at `/swimlane/`, not `/` (see `base` in `vite.config.ts`), but the code was
building the redirect URI from `window.location.origin` alone, which drops
that path. Fixed in `authConfig.ts` to append Vite's own `BASE_URL`.

**The earlier guidance below (origin only, no path) was wrong** - the
redirect URI registered in Entra has to be the exact URL the app sends,
including the path, or Microsoft rejects it (AADSTS50011 mismatch):

- `http://localhost:5173/` (local dev)
- `https://sandmansss45.github.io/swimlane/` (live site - **note the
  trailing `/swimlane/` path**, replacing whatever origin-only entry may
  already be registered)

Needs updating in Entra admin center → App registrations → Swimlane Studio
→ Authentication → Redirect URIs.

---

**Until this is done**, the app still runs and is fully clickable using the
"Use mock data (no sign-in)" button on the sign-in screen — that's fixture
data, not the real SharePoint site, so all the layout/arrow/edit work can
keep being reviewed without waiting on any of this.
