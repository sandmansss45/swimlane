import { useState, useMemo } from 'react';
import { IPublicClientApplication, InteractionStatus } from '@azure/msal-browser';
import { MsalProvider, AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from '@azure/msal-react';
import { ThemeProvider, PrimaryButton, MessageBar, MessageBarType, Spinner } from '@fluentui/react';
import SwimlaneStudio from './swimlane/components/SwimlaneStudio';
import { MockDataService } from './swimlane/services/MockDataService';
import { GraphDataService } from './swimlane/services/GraphDataService';
import { GRAPH_SCOPES, ENTRA_CLIENT_ID } from './swimlane/auth/authConfig';
import { swimlaneTheme, signInTheme } from './swimlane/theme';
import qleLogo from './assets/qle-logo.svg';
import styles from './App.module.scss';

interface IAppProps {
  msalInstance: IPublicClientApplication;
}

// Sign-in gate: real usage talks to the live SharePoint lists via Graph
// and needs a signed-in account, but there's no reason local development
// or a quick look at the UI should be blocked on the Entra app
// registration being finished - "Use mock data" skips auth entirely and
// renders against the same fixture data used during the SPFx build.
const SignedInApp: React.FC = () => {
  const { instance } = useMsal();
  const dataService = useMemo(() => new GraphDataService(instance), [instance]);
  return (
    <SwimlaneStudio
      dataService={dataService}
      signOutLabel="Sign out"
      onSignOut={() => instance.logoutRedirect()}
    />
  );
};

// AuthenticatedTemplate/UnauthenticatedTemplate both render nothing while
// MSAL is still figuring out the auth state (inProgress !== 'none') - most
// visibly right after landing back from the Microsoft sign-in redirect,
// exactly the moment someone's watching for "did that work?" - which
// without this was a blank white flash instead of any feedback at all.
const AuthStatusGate: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const { inProgress } = useMsal();
  if (inProgress !== InteractionStatus.None) {
    return (
      <ThemeProvider theme={signInTheme} className={styles.signInScreen}>
        <div className={styles.card}>
          <img src={qleLogo} className={styles.mark} alt="Quantum Leap Energy" />
          <h2 className={styles.title}>Swimlane Studio</h2>
          <Spinner label="Signing you in..." styles={{ label: { color: '#9fb0cc' } }} />
        </div>
      </ThemeProvider>
    );
  }
  // eslint-disable-next-line react/jsx-no-useless-fragment
  return <>{children}</>;
};

const SignInGate: React.FC<{ onUseMock: () => void }> = ({ onUseMock }) => {
  const { instance } = useMsal();
  const registrationPending = ENTRA_CLIENT_ID.indexOf('TODO') === 0;

  return (
    <ThemeProvider theme={signInTheme} className={styles.signInScreen}>
      <div className={styles.card}>
        <img src={qleLogo} className={styles.mark} alt="Quantum Leap Energy" />
        <h2 className={styles.title}>Swimlane Studio</h2>
        <p className={styles.subtitle}>Finance process visualization</p>

        {registrationPending && (
          <div className={styles.warning}>
            <MessageBar messageBarType={MessageBarType.warning}>
              The Entra app registration hasn&apos;t been created yet, so sign-in will fail until
              ENTRA_CLIENT_ID in src/swimlane/auth/authConfig.ts is filled in. Use mock data
              below in the meantime.
            </MessageBar>
          </div>
        )}

        <div className={styles.actions}>
          <PrimaryButton
            text="Sign in with Microsoft"
            // loginRedirect, not loginPopup - the redirect URI points at
            // this same full app bundle (there's no separate minimal
            // "blank" page for a popup to land on), so a popup ends up
            // booting the entire app inside itself instead of just
            // relaying the result back to the opener and closing, which
            // is exactly what got stuck showing a blank popup window when
            // tested for real. loginRedirect avoids the whole class of
            // bug - the same tab navigates away and back, and
            // handleRedirectPromise() in msalInstance.ts already picks up
            // the result on the next load.
            onClick={() => instance.loginRedirect({ scopes: GRAPH_SCOPES })}
          />
          <button type="button" className={styles.mockLink} onClick={onUseMock}>
            Use mock data (no sign-in)
          </button>
        </div>
      </div>
    </ThemeProvider>
  );
};

function App({ msalInstance }: IAppProps) {
  const [useMock, setUseMock] = useState(false);
  const mockService = useMemo(() => new MockDataService(), []);

  return (
    <ThemeProvider theme={swimlaneTheme}>
      {useMock ? (
        <SwimlaneStudio
          dataService={mockService}
          signOutLabel="Back to sign-in"
          onSignOut={() => setUseMock(false)}
        />
      ) : (
        <MsalProvider instance={msalInstance}>
          <AuthStatusGate>
            <AuthenticatedTemplate>
              <SignedInApp />
            </AuthenticatedTemplate>
            <UnauthenticatedTemplate>
              <SignInGate onUseMock={() => setUseMock(true)} />
            </UnauthenticatedTemplate>
          </AuthStatusGate>
        </MsalProvider>
      )}
    </ThemeProvider>
  );
}

export default App;
