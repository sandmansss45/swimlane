// Whether a swimlane (one Progress ID's flow, in one region) is still
// being worked on or considered done - a lightweight status label,
// distinct from swimlane locking (IProgressIdLock): locking blocks
// further edits and is an append-only audit trail of every lock/unlock;
// this is just a current, mutable "where is this at" marker with no
// enforcement behind it at all, same "social signal, not a technical
// barrier" philosophy the lock feature itself already uses (see
// README-HANDOVER.md). A swimlane with no record here yet is treated as
// Draft by default - nobody's explicitly marked it either way.
export type SwimlaneStage = 'Draft' | 'Finalised';

export interface ISwimlaneStatus {
  id: string;
  progressId: string;
  region: string; // '' for no specific region, same convention as IProgressIdLock
  stage: SwimlaneStage;
  // Whoever most recently toggled it - refreshed every time the stage
  // changes (unlike a lock's own createdBy, this is a single mutable
  // record, not one row per change), so it always answers "who set the
  // CURRENT stage", not "who first created this record".
  setBy: string;
  setAt: string; // ISO date string
}
