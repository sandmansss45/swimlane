import { Configuration } from '@azure/msal-browser';

// Real Application (client) ID from the "Swimlane Studio" app registration
// in the Quantum Leap Energy tenant.
export const ENTRA_CLIENT_ID = 'b654eeeb-7451-49cf-94e4-d089592de6e5';

// Real tenant directory ID, confirmed from an actual AADSTS error response
// during setup (not guessed) - more robust than the domain-name form if
// the tenant's verified domains ever change.
export const ENTRA_TENANT = 'f75d6062-398e-493a-9af8-5b532b88d594';

export const msalConfig: Configuration = {
  auth: {
    clientId: ENTRA_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${ENTRA_TENANT}`,
    // window.location.origin alone drops the /swimlane/ base path GitHub
    // Pages serves this project repo under (see vite.config.ts) - Microsoft
    // would send the signed-in user back to the bare domain instead of the
    // app itself (a real 404 this produced once tested for real).
    // import.meta.env.BASE_URL is Vite's own record of that same base path
    // (already '/' in local dev, '/swimlane/' in the GitHub Pages build),
    // so this always matches wherever the app is actually being served.
    redirectUri: window.location.origin + import.meta.env.BASE_URL
  },
  cache: {
    cacheLocation: 'localStorage'
  }
};

// CONFIRMED approach (per Method Group, 2026-08): a dedicated new
// SharePoint site named "Swimlane Studio" is being created specifically
// for this app, and the app registration's Graph permission is
// Sites.Selected rather than Sites.ReadWrite.All - this app can only ever
// touch that one site, never QLEFinance or anything else in the tenant.
// Admin consent alone is NOT enough for Sites.Selected - a tenant admin
// also has to run one additional Graph call granting this specific app
// access to this specific site once it exists (see AUTH_SETUP.md).
export const GRAPH_SCOPES = ['User.Read', 'Sites.Selected'];

export const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';

// SharePoint site identified by hostname + server-relative path, resolved
// to a real Graph site ID at runtime (see GraphDataService) rather than
// hardcoding a guessed GUID. Confirmed 2026-08-13 from the real site URL
// (qleapenergy.sharepoint.com/sites/SwimlaneStudio) once Method Group
// created it.
export const SHAREPOINT_SITE_HOSTNAME = 'qleapenergy.sharepoint.com';
export const SHAREPOINT_SITE_PATH = '/sites/SwimlaneStudio';
