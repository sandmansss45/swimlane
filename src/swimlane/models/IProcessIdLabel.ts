// A user-created name for a Process ID that doesn't have any real steps
// yet - normally a Process ID gets its name for free from a real step's
// processDescription (see getLabel in SwimlaneStudio.tsx) or from the
// static APQC_PROCESS_ID_NAMES table, but neither exists for a brand new
// Process ID someone creates as an empty shell (see "+ Add new progress
// ID") before adding any actual steps to it - without this, it would show
// the generic "Process ID X.Y.Z" fallback until a first step existed.
export interface IProcessIdLabel {
  id: string;
  processId: string; // e.g. '13.2.5'
  name: string; // e.g. 'Manage portfolio project'
}
