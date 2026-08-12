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
    redirectUri: window.location.origin
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
// hardcoding a guessed GUID.
// TODO - CONFIRM the real path once the new "Swimlane Studio" site
// actually exists - Method Group hasn't created it yet, so this is not a
// real value, just a placeholder guess at the likely URL slug.
export const SHAREPOINT_SITE_HOSTNAME = 'qleapenergy.sharepoint.com';
export const SHAREPOINT_SITE_PATH = '/sites/TODO-CONFIRM-SWIMLANE-STUDIO-SITE-PATH';
