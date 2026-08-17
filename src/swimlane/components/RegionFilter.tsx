import * as React from 'react';
import styles from './RegionFilter.module.scss';

export interface IRegionFilterProps {
  // Distinct values present in the employee list's "region" field - as of
  // 2026-08-17 this is actually sourced from Department (e.g. "3101 -
  // Security and Safety"), since the real "QLE Existing Organisation" list
  // has no dedicated region column. Read from the live employee data
  // rather than hardcoded, so this adapts to whatever values are really
  // there either way.
  regions: string[];
  selectedRegion: string | undefined; // undefined = "All"
  onChange: (region: string | undefined) => void;
}

// Narrows the "Responsible" job-title picker (see ProcessStepForm) to one
// region/department at a time - confirmed design rule (see the regions
// prop comment above for what "region" actually means for the real data).
// Same pill styling as ProcessStepTabs, not Fluent's default Pivot look -
// this row sits directly under that one in the toolbar, so a mismatched
// tab style there read as an unrelated, unfinished-looking component.
const RegionFilter: React.FC<IRegionFilterProps> = ({ regions, selectedRegion, onChange }) => {
  if (regions.length === 0) return null;

  return (
    <div className={styles.row}>
      <span className={styles.label}>Region</span>
      <div className={styles.tabs}>
        <button
          type="button"
          className={`${styles.tab} ${selectedRegion === undefined ? styles.active : ''}`}
          onClick={() => onChange(undefined)}
        >
          All
        </button>
        {regions.map(region => (
          <button
            type="button"
            key={region}
            className={`${styles.tab} ${selectedRegion === region ? styles.active : ''}`}
            onClick={() => onChange(region)}
          >
            {region}
          </button>
        ))}
      </div>
    </div>
  );
};

export default RegionFilter;
