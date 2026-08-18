import * as React from 'react';
import { IProcessStep } from '../models/IProcessStep';
import { IProgressIdLock } from '../models/IProgressIdLock';
import styles from './AuditView.module.scss';

export interface IAuditViewProps {
  steps: IProcessStep[];
  progressIdLocks: IProgressIdLock[];
}

const formatDate = (iso: string | undefined): string => (iso ? new Date(iso).toLocaleString() : '—');

// A dedicated place to review both halves of this app's audit trail in
// one screen, rather than one step or one lock banner at a time - added
// alongside per-step createdBy/modifiedBy tracking, directly in response
// to a real compliance ask ("who changed a workflow, when, why, who
// approved it... future SOX-ready"). Read-only by design - nothing here
// is ever edited from this view, it's a report, not a form.
const AuditView: React.FC<IAuditViewProps> = ({ steps, progressIdLocks }) => {
  // Most-recently-touched first - what changed lately is almost always
  // what an audit review actually wants to see, not alphabetical/ID
  // order. Steps that have never been touched by this feature (no
  // modifiedAt at all - e.g. original seed/import data predating it)
  // sort to the end rather than the top.
  const sortedSteps = React.useMemo(
    () => steps.slice().sort((a, b) => (b.modifiedAt || '').localeCompare(a.modifiedAt || '')),
    [steps]
  );

  const sortedLocks = React.useMemo(
    () => progressIdLocks.slice().sort((a, b) => (b.lockedAt || '').localeCompare(a.lockedAt || '')),
    [progressIdLocks]
  );

  return (
    <>
      <div className={styles.card}>
        <h3 className={styles.title}>Step history ({steps.length})</h3>
        <div className={styles.tableWrap}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th>Process Step ID</th>
                <th>Description</th>
                <th>Created by</th>
                <th>Created</th>
                <th>Last modified by</th>
                <th>Last modified</th>
              </tr>
            </thead>
            <tbody>
              {sortedSteps.map(s => (
                <tr key={s.id}>
                  <td>{s.processStepId}</td>
                  <td>{s.actionDescription}</td>
                  <td>{s.createdBy || <span className={styles.muted}>—</span>}</td>
                  <td className={styles.muted}>{formatDate(s.createdAt)}</td>
                  <td>{s.modifiedBy || <span className={styles.muted}>—</span>}</td>
                  <td className={styles.muted}>{formatDate(s.modifiedAt)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <div className={styles.card}>
        <h3 className={styles.title}>Swimlane locks ({progressIdLocks.length})</h3>
        {progressIdLocks.length === 0 ? (
          <p className={styles.empty}>No swimlane has ever been locked yet.</p>
        ) : (
          <div className={styles.tableWrap}>
            <table className={styles.table}>
              <thead>
                <tr>
                  <th>Progress ID</th>
                  <th>Region</th>
                  <th>Locked by</th>
                  <th>Locked at</th>
                  <th>Reason</th>
                  <th>Unlocked by</th>
                  <th>Unlocked at</th>
                  <th>Unlock reason</th>
                </tr>
              </thead>
              <tbody>
                {sortedLocks.map(l => (
                  <tr key={l.id}>
                    <td>{l.progressId}</td>
                    <td>{l.region || <span className={styles.muted}>All</span>}</td>
                    <td>{l.lockedBy}</td>
                    <td className={styles.muted}>{formatDate(l.lockedAt)}</td>
                    <td>{l.reason || <span className={styles.muted}>—</span>}</td>
                    <td>{l.unlockedBy || <span className={styles.muted}>Still locked</span>}</td>
                    <td className={styles.muted}>{formatDate(l.unlockedAt)}</td>
                    <td>{l.unlockReason || <span className={styles.muted}>—</span>}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </>
  );
};

export default AuditView;
