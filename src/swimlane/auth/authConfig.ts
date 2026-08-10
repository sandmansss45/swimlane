import { Configuration } from '@azure/msal-browser';

// TODO - confirm the real Entra Application (client) ID once the app
// registration is created (see AUTH_SETUP.md at the repo root for the
// exact steps someone with Entra admin rights needs to run). Nothing here
// will authenticate until that's filled in.
export const ENTRA_CLIENT_ID = 'TODO-CONFIRM-ENTRA-CLIENT-ID';

// The site's own hostname/domain (qleapenergy.sharepoint.com) confirms the
// tenant's primary domain is qleapenergy.com, so the authority can target
// it by verified domain name rather than a raw tenant GUID. Swap for the
// real tenant ID from the Entra admin center if this domain-based form
// ever stops resolving.
export const ENTRA_TENANT = 'qleapenergy.com';

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
