import { PublicClientApplication, EventType } from '@azure/msal-browser';
import { msalConfig } from './authConfig';

export const msalInstance = new PublicClientApplication(msalConfig);

export async function initMsal(): Promise<void> {
  await msalInstance.initialize();
  const response = await msalInstance.handleRedirectPromise();
  if (response?.account) {
    msalInstance.setActiveAccount(response.account);
  } else if (!msalInstance.getActiveAccount() && msalInstance.getAllAccounts().length > 0) {
    msalInstance.setActiveAccount(msalInstance.getAllAccounts()[0]);
  }

  msalInstance.addEventCallback(event => {
    if (event.eventType === EventType.LOGIN_SUCCESS && event.payload) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const account = (event.payload as any).account;
      if (account) msalInstance.setActiveAccount(account);
    }
  });
}
