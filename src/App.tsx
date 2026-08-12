import { useState, useMemo } from 'react';
import { IPublicClientApplication } from '@azure/msal-browser';
import { MsalProvider, AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from '@azure/msal-react';
import { ThemeProvider, PrimaryButton, MessageBar, MessageBarType } from '@fluentui/react';
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
  return <SwimlaneStudio dataService={dataService} />;
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
            onClick={() => instance.loginPopup({ scopes: GRAPH_SCOPES })}
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
        <SwimlaneStudio dataService={mockService} />
      ) : (
        <MsalProvider instance={msalInstance}>
          <AuthenticatedTemplate>
            <SignedInApp />
          </AuthenticatedTemplate>
          <UnauthenticatedTemplate>
            <SignInGate onUseMock={() => setUseMock(true)} />
          </UnauthenticatedTemplate>
        </MsalProvider>
      )}
    </ThemeProvider>
  );
}

export default App;
