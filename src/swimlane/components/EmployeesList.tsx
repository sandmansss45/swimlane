import * as React from 'react';
import { IEmployee } from '../models/IEmployee';
import styles from './EmployeesList.module.scss';

export interface IEmployeesListProps {
  employees: IEmployee[];
}

// No region/department filter here (deliberately removed) - the real
// Employees list has dozens of granular departments (e.g. "3602 - Quantum
// Enrichment"), which as filter tabs overflowed the header entirely and
// made the tab bar unusable. The swimlane's own region filter (narrowing
// the Responsible picker for one step) is unaffected - this only touches
// the standalone directory view.
const EmployeesList: React.FC<IEmployeesListProps> = ({ employees }) => {
  const sorted = React.useMemo(() => {
    // Employee names aren't pulled into the app at all (explicit user
    // choice - see models/IEmployee.ts), so several people holding the
    // same title in the same department would otherwise render as
    // identical-looking duplicate rows. Dedupe down to one row per
    // distinct title+region pair instead.
    const seen = new Set<string>();
    const distinct = employees.filter(e => {
      const key = `${e.jobTitle}|${e.region || ''}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
    return distinct.sort((a, b) => a.jobTitle.localeCompare(b.jobTitle));
  }, [employees]);

  return (
    <div className={styles.card}>
      <div className={styles.toolbar}>
        <h3 className={styles.title}>Employees ({sorted.length})</h3>
      </div>

      {sorted.length === 0 ? (
        <p className={styles.empty}>No employees loaded yet.</p>
      ) : (
        <div className={styles.tableWrap}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th>Job Title</th>
                <th>Region</th>
              </tr>
            </thead>
            <tbody>
              {sorted.map(e => (
                <tr key={e.id}>
                  <td>{e.jobTitle || <span className={styles.muted}>—</span>}</td>
                  <td>{e.region || <span className={styles.muted}>—</span>}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
};

export default EmployeesList;
