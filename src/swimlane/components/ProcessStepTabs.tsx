import * as React from 'react';
import { IProcessStep } from '../models/IProcessStep';
import { compareProcessStepIds } from '../utils/columns';
import styles from './ProcessStepTabs.module.scss';

export interface IProcessStepTabsProps {
  steps: IProcessStep[]; // all steps for the currently selected Process ID
  selectedStepId: string | undefined; // undefined = "All" (top-level view)
  onSelect: (stepId: string | undefined) => void;
  // Opens the lightweight section-creation flow (see AddStepSectionModal)
  // - undefined when there's nowhere sensible to add one yet (no Progress
  // ID selected).
  onAddNew?: () => void;
}

// Drill-down level of the two confirmed navigation levels: "All" shows
// every row under the Process ID as one continuous flow; each tab
// narrows to a single Process Step ID.
const ProcessStepTabs: React.FC<IProcessStepTabsProps> = ({ steps, selectedStepId, onSelect, onAddNew }) => {
  const stepIds = React.useMemo(
    () => Array.from(new Set(steps.map(s => s.processStepId))).sort(compareProcessStepIds),
    [steps]
  );

  const countFor = (stepId: string | undefined): number =>
    stepId === undefined ? steps.length : steps.filter(s => s.processStepId === stepId).length;

  return (
    <div className={styles.tabs}>
      <button
        className={`${styles.tab} ${selectedStepId === undefined ? styles.active : ''}`}
        onClick={() => onSelect(undefined)}
      >
        All <span className={styles.count}>{countFor(undefined)}</span>
      </button>
      {stepIds.map(stepId => (
        <button
          key={stepId}
          className={`${styles.tab} ${selectedStepId === stepId ? styles.active : ''}`}
          onClick={() => onSelect(stepId)}
        >
          {stepId} <span className={styles.count}>{countFor(stepId)}</span>
        </button>
      ))}
      {onAddNew && (
        <button type="button" className={styles.addTab} onClick={onAddNew}>
          + Add section
        </button>
      )}
    </div>
  );
};

export default ProcessStepTabs;
