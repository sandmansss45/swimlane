import * as React from 'react';
import { IProcessStep, getProgressId } from '../models/IProcessStep';
import { compareProcessStepIds } from '../utils/columns';
import styles from './ProgressIdPicker.module.scss';

export interface IProgressIdPickerProps {
  steps: IProcessStep[];
  onSelect: (progressId: string) => void;
}

// Entry point above the two confirmed navigation levels: swimlanes group
// by Progress ID (e.g. 9.6.1 contains 9.6.1.1 through 9.6.1.6 as one
// continuous flow), so pick which Progress ID to open before showing its
// top-level flow / drill-down views.
const ProgressIdPicker: React.FC<IProgressIdPickerProps> = ({ steps, onSelect }) => {
  const groups = React.useMemo(() => {
    const byProgressId = new Map<string, { progressId: string; description: string; count: number }>();
    steps.forEach(step => {
      const progressId = getProgressId(step.processStepId);
      const existing = byProgressId.get(progressId);
      if (existing) {
        existing.count += 1;
      } else {
        byProgressId.set(progressId, { progressId, description: step.processDescription, count: 1 });
      }
    });
    return Array.from(byProgressId.values()).sort((a, b) => compareProcessStepIds(a.progressId, b.progressId));
  }, [steps]);

  return (
    <div className={styles.grid}>
      {groups.map(group => (
        <div className={styles.card} key={group.progressId} onClick={() => onSelect(group.progressId)}>
          <span className={styles.code}>{group.progressId}</span>
          <h3>{group.description}</h3>
          <p>{group.count} step{group.count === 1 ? '' : 's'}</p>
        </div>
      ))}
      {groups.length === 0 && <p>No process data loaded yet.</p>}
    </div>
  );
};

export default ProgressIdPicker;
