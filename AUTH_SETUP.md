# One-time setup: Entra app registration

The app can't sign anyone in or call Microsoft Graph until this exists. It's a
one-time step in the Entra admin center — not something I can do from here,
and the "grant admin consent" step needs whoever administers your Entra ID
tenant (may or may not be the same people as "Method Group" for your local
machine — worth checking).

## 1. Create the registration

1. Go to https://entra.microsoft.com (or portal.azure.com → Microsoft Entra ID) → **App registrations** → **New registration**.
2. Name: `Swimlane Studio` (anything recognizable is fine).
3. Supported account types: **Accounts in this organizational directory only** (single tenant).
4. Redirect URI: platform = **Single-page application (SPA)**, URI = `http://localhost:5173`.
5. Click **Register**.
6. Back on **Authentication**, add a second SPA redirect URI once GitHub Pages
   is live: `https://<your-github-username>.github.io` (origin only, no
   `/repo-name/` path — the app matches on origin, not full path).

## 2. Copy the client ID

On the registration's **Overview** page, copy the **Application (client) ID**
(a GUID) and send it back — I'll drop it into
`src/swimlane/auth/authConfig.ts` as `ENTRA_CLIENT_ID`.

## 3. Add Graph permissions

1. Left nav → **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**.
2. Add `Sites.ReadWrite.All` (broad, simplest to consent). `User.Read` is usually
   there by default.
3. Click **Grant admin consent for [tenant name]** — this is the button that
   needs an Entra admin, not just any user. Nothing works until this is green-checked.

   **Least-privilege alternative:** instead of `Sites.ReadWrite.All` (every
   site in the tenant), an admin can grant `Sites.Selected` and then run one
   Graph/PowerShell command to scope this app to just the QLEFinance site.
   More steps, but the app literally cannot see any other site's data. Say
   the word if you'd rather do it this way and I'll write out those exact
   commands.

## 4. Tell me the client ID (and which permission model you used)

Once steps 1–3 are done, send me the GUID. I'll wire it in and we can test
real sign-in against `http://localhost:5173`.

---

**Until this is done**, the app still runs and is fully clickable using the
"Use mock data (no sign-in)" button on the sign-in screen — that's the same
fixture data used throughout the SPFx build, so all the layout/arrow/edit
work can keep being reviewed without waiting on this.
