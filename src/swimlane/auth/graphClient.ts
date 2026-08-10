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

  constructor(msal: IPublicClientApplication) {
    this.msal = msal;
  }

  private async getToken(): Promise<string> {
    const account = this.msal.getActiveAccount() || this.msal.getAllAccounts()[0];
    if (!account) {
      throw new Error('No signed-in account - sign in before calling Microsoft Graph.');
    }
    try {
      const result = await this.msal.acquireTokenSilent({ scopes: GRAPH_SCOPES, account });
      return result.accessToken;
    } catch {
      const result = await this.msal.acquireTokenPopup({ scopes: GRAPH_SCOPES, account });
      return result.accessToken;
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
