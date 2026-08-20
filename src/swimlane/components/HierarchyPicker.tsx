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
  // wired in at all three levels (Category, Process Group, Progress ID),
  // each of which can carry a user-assigned custom label that wins over
  // whatever name would otherwise show (static APQC table, or a step's
  // own derived text) - see customCategoryNames/customGroupNames/
  // customProgressIdNames in SwimlaneStudio.tsx. Renaming one of the
  // fixed 13 real APQC categories creates a new custom label the first
  // time (same "first rename creates a label" behaviour as the other two
  // levels), rather than editing anything in the static table itself.
  onRename?: (groupId: string, currentLabel: string) => void;
  // When set, a card's count shows the number of DISTINCT values this
  // returns per step (e.g. distinct Process Group IDs), not the number of
  // raw process steps - used at the Category level, where "steps" was a
  // less useful number than "how many process groups this category has".
  // Process Group and Progress ID levels leave this unset and keep
  // showing their own step counts.
  countBy?: (processStepId: string) => string;
  countLabel?: string; // singular noun for the count, e.g. "process group" - defaults to "step"
  // Plural form, for nouns a trailing "s" doesn't work for (e.g. "process"
  // -> "processes", not "processs") - defaults to countLabel + "s" when
  // that's actually correct (e.g. "step"/"process group").
  countLabelPlural?: string;
  // The full static universe of countBy's sub-group IDs (e.g. every known
  // Process Group ID from APQC_PROCESS_GROUP_NAMES), same reasoning as
  // allGroupIds above but one level deeper - without this, a category
  // with zero steps loaded yet (all its process groups still empty)
  // showed "0 process groups" even though the APQC standard defines
  // several for it, since counting only ran over steps that actually
  // exist. Ignored unless countBy is also set.
  allSubGroupIds?: string[];
}

// Generic drill-down level: groups steps by whatever ID prefix the caller
// asks for (Category, Process Group, or Progress ID - see the APQC
// hierarchy comment in utils/apqcHierarchy.ts) and shows one card per
// group. The same component powers all three levels above the swimlane
// itself, so they look and behave identically by construction.
const HierarchyPicker: React.FC<IHierarchyPickerProps> = ({
  steps, getGroupId, getLabel, onSelect, emptyMessage, allGroupIds, onAddNew, addNewLabel, onRename,
  countBy, countLabel, countLabelPlural, allSubGroupIds
}) => {
  const groups = React.useMemo(() => {
    const byGroupId = new Map<string, { groupId: string; label: string; count: number; subIds: Set<string> }>();
    steps.forEach(step => {
      const groupId = getGroupId(step.processStepId);
      const subId = countBy ? countBy(step.processStepId) : undefined;
      const existing = byGroupId.get(groupId);
      if (existing) {
        existing.count += 1;
        if (subId) existing.subIds.add(subId);
      } else {
        const subIds = new Set<string>();
        if (subId) subIds.add(subId);
        byGroupId.set(groupId, { groupId, label: getLabel(groupId, step), count: 1, subIds });
      }
    });
    (allGroupIds || []).forEach(groupId => {
      if (!byGroupId.has(groupId)) byGroupId.set(groupId, { groupId, label: getLabel(groupId), count: 0, subIds: new Set() });
    });
    // Same static-universe fallback as allGroupIds, one level deeper - a
    // sub-group ID that's part of the APQC standard counts toward its
    // parent's card even if no step has landed in it yet.
    if (countBy) {
      (allSubGroupIds || []).forEach(subId => {
        const groupId = getGroupId(subId);
        const existing = byGroupId.get(groupId);
        if (existing) {
          existing.subIds.add(subId);
        } else {
          byGroupId.set(groupId, { groupId, label: getLabel(groupId), count: 0, subIds: new Set([subId]) });
        }
      });
    }
    return Array.from(byGroupId.values()).sort((a, b) => compareProcessStepIds(a.groupId, b.groupId));
  }, [steps, getGroupId, getLabel, allGroupIds, countBy, allSubGroupIds]);

  return (
    <div className={styles.grid}>
      {groups.map(group => {
        const displayCount = countBy ? group.subIds.size : group.count;
        const singular = countLabel || 'step';
        const plural = countLabelPlural || `${singular}s`;
        return (
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
            <p>{displayCount} {displayCount === 1 ? singular : plural}</p>
          </div>
        );
      })}
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
