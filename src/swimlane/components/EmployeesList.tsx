import * as React from 'react';
import { DefaultButton } from '@fluentui/react';
import { IEmployee } from '../models/IEmployee';
import styles from './EmployeesList.module.scss';

export interface IEmployeesListProps {
  employees: IEmployee[];
  onAddClick: () => void;
}

// No department filter here (deliberately removed) - the real Employees
// list has dozens of granular departments (e.g. "3602 - Quantum
// Enrichment"), which as filter tabs overflowed the header entirely and
// made the tab bar unusable. The swimlane's own department filter
// (narrowing the Responsible picker for one step) is unaffected - this
// only touches the standalone directory view.
const EmployeesList: React.FC<IEmployeesListProps> = ({ employees, onAddClick }) => {
  const [query, setQuery] = React.useState('');

  const sorted = React.useMemo(() => {
    // Employee names aren't pulled into the app at all (explicit user
    // choice - see models/IEmployee.ts), so several people holding the
    // same title in the same department would otherwise render as
    // identical-looking duplicate rows. Dedupe down to one row per
    // distinct title+department pair instead.
    const seen = new Set<string>();
    const distinct = employees.filter(e => {
      const key = `${e.jobTitle}|${e.department || ''}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
    return distinct.sort((a, b) => a.jobTitle.localeCompare(b.jobTitle));
  }, [employees]);

  // A plain scroll-and-squint list stopped being enough once the real
  // directory grew past a couple dozen distinct job titles - matches
  // either column so "quantum" finds both a job title and a department
  // without needing to know which one the word lives in.
  const filtered = React.useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return sorted;
    return sorted.filter(e => e.jobTitle.toLowerCase().includes(q) || (e.department || '').toLowerCase().includes(q));
  }, [sorted, query]);

  return (
    <div className={styles.card}>
      <div className={styles.toolbar}>
        <h3 className={styles.title}>Employees ({sorted.length})</h3>
        <input
          className={styles.searchInput}
          type="text"
          placeholder="Search job title or department..."
          value={query}
          onChange={e => setQuery(e.target.value)}
        />
        <DefaultButton text="+ Add employee" iconProps={{ iconName: 'AddFriend' }} onClick={onAddClick} />
      </div>

      {sorted.length === 0 ? (
        <p className={styles.empty}>No employees loaded yet.</p>
      ) : filtered.length === 0 ? (
        <p className={styles.empty}>No job titles or departments match "{query}".</p>
      ) : (
        <div className={styles.tableWrap}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th>Job Title</th>
                <th>Department</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map(e => (
                <tr key={e.id}>
                  <td>{e.jobTitle || <span className={styles.muted}>—</span>}</td>
                  <td>{e.department || <span className={styles.muted}>—</span>}</td>
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
