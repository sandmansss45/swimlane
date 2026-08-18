import * as React from 'react';
import { IProcessStep, KNOWN_FLOW_REGIONS } from '../models/IProcessStep';
import styles from './FlowRegionTabs.module.scss';

export interface IFlowRegionTabsProps {
  steps: IProcessStep[]; // all steps for the currently selected Progress ID, UNFILTERED by region - the tab list needs to see every region present, not just the selected one
  selectedRegion: string | undefined; // undefined = "All"
  onSelect: (region: string | undefined) => void;
}

// A genuinely different concept from ProcessStepTabs just below it in the
// toolbar: those narrow within ONE continuous flow, this picks between
// entirely separate swimlanes for the same Progress ID (e.g. the UK AP
// process vs the US AP process) - confirmed design rule, added at a real
// user's request after Department (an unrelated per-employee grouping)
// turned out not to cover this at all. Sits above ProcessStepTabs since
// it's the bigger partition - which region you're in determines which
// Process Step ID tabs and steps even show.
const FlowRegionTabs: React.FC<IFlowRegionTabsProps> = ({ steps, selectedRegion, onSelect }) => {
  const regions = React.useMemo(() => {
    const fromData = Array.from(new Set(steps.map(s => s.region).filter((r): r is string => !!r)));
    const extras = fromData.filter(r => !KNOWN_FLOW_REGIONS.includes(r)).sort();
    return [...KNOWN_FLOW_REGIONS, ...extras];
  }, [steps]);

  const countFor = (region: string | undefined): number =>
    region === undefined ? steps.length : steps.filter(s => s.region === region).length;

  return (
    <div className={styles.tabs}>
      <button
        type="button"
        className={`${styles.tab} ${selectedRegion === undefined ? styles.active : ''}`}
        onClick={() => onSelect(undefined)}
      >
        All <span className={styles.count}>{countFor(undefined)}</span>
      </button>
      {regions.map(region => (
        <button
          type="button"
          key={region}
          className={`${styles.tab} ${selectedRegion === region ? styles.active : ''}`}
          onClick={() => onSelect(region)}
        >
          {region} <span className={styles.count}>{countFor(region)}</span>
        </button>
      ))}
    </div>
  );
};

export default FlowRegionTabs;
