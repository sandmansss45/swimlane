// A user-created name for a Progress ID that doesn't have any real steps
// yet - normally a Progress ID gets its name for free from a real step's
// processDescription (see getLabel in SwimlaneStudio.tsx) or from the
// static APQC_PROGRESS_ID_NAMES table, but neither exists for a brand new
// Progress ID someone creates as an empty shell (see "+ Add new progress
// ID") before adding any actual steps to it - without this, it would show
// the generic "Progress ID X.Y.Z" fallback until a first step existed.
export interface IProgressIdLabel {
  id: string;
  progressId: string; // e.g. '13.2.5'
  name: string; // e.g. 'Manage portfolio project'
}
