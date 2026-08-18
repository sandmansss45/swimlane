import * as React from 'react';
import styles from './DepartmentFilter.module.scss';

export interface IDepartmentFilterProps {
  // Distinct values present in the employee list's "department" field -
  // CONFIRMED 2026-08-17 the real "QLE Existing Organisation" list has no
  // dedicated region/geography column at all, so this is genuinely
  // department data (e.g. "3101 - Security and Safety"), not a region.
  // Read from the live employee data rather than hardcoded, so this
  // adapts to whatever values are really there.
  departments: string[];
  selectedDepartment: string | undefined; // undefined = "All"
  onChange: (department: string | undefined) => void;
}

// Narrows the "Responsible" job-title picker (see ProcessStepForm) to one
// department at a time. Same pill styling as ProcessStepTabs, not Fluent's
// default Pivot look - this row sits directly under that one in the
// toolbar, so a mismatched tab style there read as an unrelated,
// unfinished-looking component. Was called RegionFilter until a real user
// pointed out the values shown (department codes) aren't regions at all -
// renamed throughout rather than just relabeled, so the code matches what
// the data actually is.
const DepartmentFilter: React.FC<IDepartmentFilterProps> = ({ departments, selectedDepartment, onChange }) => {
  if (departments.length === 0) return null;

  return (
    <div className={styles.row}>
      <span className={styles.label}>Department</span>
      <div className={styles.tabs}>
        <button
          type="button"
          className={`${styles.tab} ${selectedDepartment === undefined ? styles.active : ''}`}
          onClick={() => onChange(undefined)}
        >
          All
        </button>
        {departments.map(department => (
          <button
            type="button"
            key={department}
            className={`${styles.tab} ${selectedDepartment === department ? styles.active : ''}`}
            onClick={() => onChange(department)}
          >
            {department}
          </button>
        ))}
      </div>
    </div>
  );
};

export default DepartmentFilter;
