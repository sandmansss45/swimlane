import * as React from 'react';
import { IProcessStep } from '../models/IProcessStep';
import { compareProcessStepIds } from '../utils/columns';
import styles from './HierarchyPicker.module.scss';

export interface IHierarchyPickerProps {
  steps: IProcessStep[];
  getGroupId: (processStepId: string) => string;
  getLabel: (groupId: string, sampleStep?: IProcessStep) => string;
  onSelect: (groupId: string) => void;
  emptyMessage?: string;
  // Group IDs that should always render a card even with zero matching
  // steps yet - e.g. the Category level always shows the full fixed set
  // of 13 APQC categories (see APQC_CATEGORY_NAMES), not just whichever
  // ones happen to have data loaded, so there's somewhere to land a brand
  // new process via "+ Add new process" instead of it only being
  // reachable once a step already exists there.
  allGroupIds?: string[];
  // Renders one more same-sized card at the end of the grid, styled to
  // stand out, that opens the "add new" flow right where the other cards
  // already are instead of only up in a separate toolbar button.
  onAddNew?: () => void;
  addNewLabel?: string;
  // Small edit affordance on each real card (not the "+ Add new" tile) -
  // only wired in at the Process Group level, since that's the only level
  // with a name that can actually be changed (Categories are the fixed
  // APQC standard; Progress ID/Activity names come from real step data,
  // not a label someone assigns).
  onRename?: (groupId: string, currentLabel: string) => void;
}

// Generic drill-down level: groups steps by whatever ID prefix the caller
// asks for (Category, Process Group, or Progress ID - see the APQC
// hierarchy comment in utils/apqcHierarchy.ts) and shows one card per
// group. The same component powers all three levels above the swimlane
// itself, so they look and behave identically by construction.
const HierarchyPicker: React.FC<IHierarchyPickerProps> = ({
  steps, getGroupId, getLabel, onSelect, emptyMessage, allGroupIds, onAddNew, addNewLabel, onRename
}) => {
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
    (allGroupIds || []).forEach(groupId => {
      if (!byGroupId.has(groupId)) byGroupId.set(groupId, { groupId, label: getLabel(groupId), count: 0 });
    });
    return Array.from(byGroupId.values()).sort((a, b) => compareProcessStepIds(a.groupId, b.groupId));
  }, [steps, getGroupId, getLabel, allGroupIds]);

  return (
    <div className={styles.grid}>
      {groups.map(group => (
        <div className={styles.card} key={group.groupId} onClick={() => onSelect(group.groupId)}>
          {onRename && (
            <button
              type="button"
              className={styles.renameButton}
              title="Rename"
              onClick={e => { e.stopPropagation(); onRename(group.groupId, group.label); }}
            >
              ✎
            </button>
          )}
          <span className={styles.code}>{group.groupId}</span>
          <h3>{group.label}</h3>
          <p>{group.count} step{group.count === 1 ? '' : 's'}</p>
        </div>
      ))}
      {onAddNew && (
        <div className={styles.addCard} onClick={onAddNew}>
          <span className={styles.addIcon}>+</span>
          <h3>{addNewLabel || 'Add new'}</h3>
        </div>
      )}
      {groups.length === 0 && !onAddNew && <p>{emptyMessage || 'No process data loaded yet.'}</p>}
    </div>
  );
};

export default HierarchyPicker;
