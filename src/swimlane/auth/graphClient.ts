import { IPublicClientApplication } from '@azure/msal-browser';
import { GRAPH_BASE_URL, GRAPH_SCOPES } from './authConfig';

/**
 * Thin fetch wrapper that acquires a Graph token silently (falling back to
 * an interactive popup only if the silent request fails - e.g. the very
 * first call in a session, or a consent/conditional-access prompt) and
 * throws with the real response body on a non-2xx status instead of
 * swallowing it, so a broken field name or missing permission surfaces as
 * a real error rather than quietly returning nothing.
 */
export class GraphClient {
  private msal: IPublicClientApplication;
  // Every list load fires several Graph calls in parallel (see
  // Promise.allSettled in SwimlaneStudio.tsx) - without this, an expired
  // silent token means each one independently races to be the one
  // interactive request MSAL allows at a time, and the losers fail
  // outright with "interaction_in_progress" instead of just waiting their
  // turn. Sharing one in-flight promise across all callers fixes that.
  private tokenPromise: Promise<string> | undefined;

  constructor(msal: IPublicClientApplication) {
    this.msal = msal;
  }

  private getToken(): Promise<string> {
    if (!this.tokenPromise) {
      this.tokenPromise = this.acquireToken().finally(() => { this.tokenPromise = undefined; });
    }
    return this.tokenPromise;
  }

  private async acquireToken(): Promise<string> {
    const account = this.msal.getActiveAccount() || this.msal.getAllAccounts()[0];
    if (!account) {
      throw new Error('No signed-in account - sign in before calling Microsoft Graph.');
    }
    try {
      const result = await this.msal.acquireTokenSilent({ scopes: GRAPH_SCOPES, account });
      return result.accessToken;
    } catch {
      // acquireTokenRedirect, not acquireTokenPopup - same reasoning as
      // the sign-in button in App.tsx: this app's redirect URI boots the
      // whole bundle rather than a minimal blank popup landing page, and
      // on top of that a popup triggered from a background data refresh
      // (not a direct click) gets killed outright by the browser's popup
      // blocker - confirmed as the actual cause of a real
      // "popup_window_error". Redirect reloads the tab instead, which the
      // existing handleRedirectPromise() bootstrap in msalInstance.ts
      // already resumes from on the next load, same as it does for the
      // initial sign-in.
      await this.msal.acquireTokenRedirect({ scopes: GRAPH_SCOPES, account });
      // Never actually reached - acquireTokenRedirect navigates the tab
      // away before its promise resolves.
      return '';
    }
  }

  public async get<T>(path: string): Promise<T> {
    const token = await this.getToken();
    const res = await fetch(`${GRAPH_BASE_URL}${path}`, {
      headers: { Authorization: `Bearer ${token}` }
    });
    if (!res.ok) {
      throw new Error(`Graph GET ${path} failed: ${res.status} ${await res.text()}`);
    }
    return res.json();
  }

  /**
   * Follows @odata.nextLink until exhausted and concatenates every page's
   * `value` array - Graph paginates list items by default (usually 200 at
   * a time), and DependsOn's row-number scheme means a partial, unsorted
   * fetch would silently break dependency resolution.
   */
  public async getAllPages<T>(path: string): Promise<T[]> {
    const token = await this.getToken();
    let url = `${GRAPH_BASE_URL}${path}`;
    const items: T[] = [];
    while (url) {
      const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
      if (!res.ok) {
        throw new Error(`Graph GET ${url} failed: ${res.status} ${await res.text()}`);
      }
      const page = await res.json();
      items.push(...(page.value || []));
      url = page['@odata.nextLink'] || '';
    }
    return items;
  }

  public async post<T>(path: string, body: unknown): Promise<T> {
    const token = await this.getToken();
    const res = await fetch(`${GRAPH_BASE_URL}${path}`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
    if (!res.ok) {
      throw new Error(`Graph POST ${path} failed: ${res.status} ${await res.text()}`);
    }
    return res.json();
  }

  public async patch(path: string, body: unknown): Promise<void> {
    const token = await this.getToken();
    const res = await fetch(`${GRAPH_BASE_URL}${path}`, {
      method: 'PATCH',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
    if (!res.ok) {
      throw new Error(`Graph PATCH ${path} failed: ${res.status} ${await res.text()}`);
    }
  }

  public async delete(path: string): Promise<void> {
    const token = await this.getToken();
    const res = await fetch(`${GRAPH_BASE_URL}${path}`, {
      method: 'DELETE',
      headers: { Authorization: `Bearer ${token}` }
    });
    if (!res.ok) {
      throw new Error(`Graph DELETE ${path} failed: ${res.status} ${await res.text()}`);
    }
  }
}
