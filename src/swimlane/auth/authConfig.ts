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

// Sites.ReadWrite.All is the broadest option and the simplest to get
// admin consent for; Sites.Selected (scoped to just the QLEFinance site)
// is the least-privilege alternative but needs an extra Graph/PowerShell
// step from an admin to grant this app access to that one site - ask
// whoever administers Entra ID which they'd rather set up.
export const GRAPH_SCOPES = ['User.Read', 'Sites.ReadWrite.All'];

export const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';

// SharePoint site identified by hostname + server-relative path, resolved
// to a real Graph site ID at runtime (see GraphDataService) rather than
// hardcoding a guessed GUID.
export const SHAREPOINT_SITE_HOSTNAME = 'qleapenergy.sharepoint.com';
export const SHAREPOINT_SITE_PATH = '/sites/QLEFinance';
