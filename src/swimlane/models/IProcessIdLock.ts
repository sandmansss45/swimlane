// A compliance/audit-trail feature (confirmed 2026-08-18, in response to a
// real "who changed this, when, why, who approved it - future SOX-ready"
// ask): locking a swimlane both freezes it against further edits AND
// leaves a permanent, attributed record that it was reviewed and signed
// off. Deliberately append-only - locking always creates a NEW record
// rather than editing an old one, so re-locking after a change leaves a
// full history (every lock/unlock cycle, who did each, when, why) instead
// of a single mutable "current status" flag that would overwrite its own
// history. A swimlane's CURRENT status is just "does its most recent
// lock record for this processId+region have a blank unlockedAt?" - see
// findActiveLock.
//
// Scoped to processId + region together, not just processId - a
// Process ID can hold several genuinely separate swimlanes side by side
// (see FlowRegionTabs), so locking the UK version of a flow shouldn't
// freeze the US/SA versions on the same Process ID. region is '' when
// locked while viewing "All" (the normal case for a Process ID that
// doesn't use region splitting at all - most of the real data today).
//
// Confirmed 2026-08-18: v1 doesn't restrict WHO can lock/unlock to
// specific people - anyone signed in can, by explicit choice. A real
// roles/permissions system is a separate, bigger feature; accountability
// here comes from every action being attributed and logged, not from a
// technical barrier.
export interface IProcessIdLock {
  id: string; // SharePoint list item ID
  processId: string; // e.g. '9.6.1'
  region: string; // '' = locked while viewing "All"
  lockedBy: string;
  lockedAt: string; // ISO date string
  reason: string; // optional context for the LOCK, '' if none given
  unlockedBy: string; // '' while still locked
  unlockedAt: string; // '' while still locked
  // Separate from `reason` above (which is why it was LOCKED) - this is
  // why it was later reopened, e.g. "correcting an error found in
  // review". Kept apart so unlocking never overwrites the original lock's
  // own reason.
  unlockReason: string;
}

/** True while a lock record is still in effect (hasn't been unlocked yet). */
export function isLockActive(lock: IProcessIdLock): boolean {
  return !lock.unlockedAt;
}

/**
 * The lock currently in effect for a given swimlane (processId + region),
 * if any - the most recently locked record that hasn't since been
 * unlocked. Undefined means this swimlane is either editable right now,
 * or has never been locked at all - both read the same way to the rest
 * of the app.
 */
export function findActiveLock(
  locks: IProcessIdLock[],
  processId: string,
  region: string | undefined
): IProcessIdLock | undefined {
  const normalizedRegion = region || '';
  const matches = locks.filter(l => l.processId === processId && l.region === normalizedRegion && isLockActive(l));
  if (matches.length === 0) return undefined;
  // Defensive: there should only ever be one active lock per swimlane
  // (the UI never offers "lock" again while one's already active), but if
  // two ever exist, the most recent one governs.
  return matches.reduce((latest, l) => (l.lockedAt > latest.lockedAt ? l : latest));
}
