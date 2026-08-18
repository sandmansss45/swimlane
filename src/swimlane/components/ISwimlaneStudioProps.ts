import { IDataService } from '../services/IDataService';

export interface ISwimlaneStudioProps {
  dataService: IDataService;
  // Only provided when there's actually somewhere to go back to (a real
  // signed-in session, or mock mode entered from the sign-in gate) - a
  // page load that lands here some other way just won't render the
  // control. Without this, a cached Microsoft session with an empty real
  // list left no way back to mock data short of clearing site storage.
  onSignOut?: () => void;
  signOutLabel?: string;
  // Attributes swimlane locks/unlocks to a real name (see
  // IProgressIdLock) - the signed-in MSAL account's display name in real
  // usage, a fixed stand-in in mock mode since there's no real session to
  // read one from.
  currentUserName: string;
}
