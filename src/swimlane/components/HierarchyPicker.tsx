import * as React from 'react';
import { IProcessStep } from '../models/IProcessStep';
import { compareProcessStepIds } from '../utils/columns';
import styles from './HierarchyPicker.module.scss';

export interface IHierarchyPickerProps {
  steps: IProcessStep[];
  getGroupId: (processStepId: string) => string;
  getLabel: (groupId: string, sampleStep: IProcessStep) => string;
  onSelect: (groupId: string) => void;
  emptyMessage?: string;
}

// Generic drill-down level: groups steps by whatever ID prefix the caller
// asks for (Category, Process Group, or Progress ID - see the APQC
// hierarchy comment in utils/apqcHierarchy.ts) and shows one card per
// group. The same component powers all three levels above the swimlane
// itself, so they look and behave identically by construction.
const HierarchyPicker: React.FC<IHierarchyPickerProps> = ({ steps, getGroupId, getLabel, onSelect, emptyMessage }) => {
  const groups = React.useMemo(() => {
    const byGroupId = new Map<string, { groupId: string; label: string; count: number }>();
    steps.forEach(step => {
      const groupId = getGroupId(step.processStepId);
      const existing = byGroupId.get(groupId);
      if (existing) {
        existing.count += 1;
      } else {
        byGroupId.set(groupId, { groupId, label: getLabel(groupId, step), count: 1 });
      }
    });
    return Array.from(byGroupId.values()).sort((a, b) => compareProcessStepIds(a.groupId, b.groupId));
  }, [steps, getGroupId, getLabel]);

  return (
    <div className={styles.grid}>
      {groups.map(group => (
        <div className={styles.card} key={group.groupId} onClick={() => onSelect(group.groupId)}>
          <span className={styles.code}>{group.groupId}</span>
          <h3>{group.label}</h3>
          <p>{group.count} step{group.count === 1 ? '' : 's'}</p>
        </div>
      ))}
      {groups.length === 0 && <p>{emptyMessage || 'No process data loaded yet.'}</p>}
    </div>
  );
};

export default HierarchyPicker;
