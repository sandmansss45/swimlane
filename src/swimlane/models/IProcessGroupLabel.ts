// A user-created name for a Process Group ID the static APQC_PROCESS_GROUP_NAMES
// lookup (apqcHierarchy.ts) doesn't already cover - e.g. a brand new "9.6.4"
// the business adds that isn't part of the confirmed reference sheet.
// Progress IDs and Activities don't need this: they get a real name for
// free from processDescription/processStepName the moment a step exists
// under them. A Process Group has no such per-step field to fall back on,
// so without this it would show the generic "Process Group 9.6.4" forever.
export interface IProcessGroupLabel {
  id: string;
  groupId: string; // e.g. '9.6.4'
  name: string; // e.g. 'Manage petty cash'
}
