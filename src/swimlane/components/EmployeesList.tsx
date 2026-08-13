import * as React from 'react';
import { IEmployee } from '../models/IEmployee';
import RegionFilter from './RegionFilter';
import styles from './EmployeesList.module.scss';

export interface IEmployeesListProps {
  employees: IEmployee[];
}

const EmployeesList: React.FC<IEmployeesListProps> = ({ employees }) => {
  const [selectedRegion, setSelectedRegion] = React.useState<string | undefined>(undefined);

  const regions = React.useMemo(
    () => Array.from(new Set(employees.map(e => e.region).filter((r): r is string => !!r))),
    [employees]
  );

  const sorted = React.useMemo(() => {
    const filtered = selectedRegion ? employees.filter(e => e.region === selectedRegion) : employees;
    // Grouped by job title first, same as how lanes are always job titles
    // - lets someone scan straight to "who else holds this title" without
    // hunting through an alphabetical-by-name list.
    return [...filtered].sort((a, b) => a.jobTitle.localeCompare(b.jobTitle) || a.name.localeCompare(b.name));
  }, [employees, selectedRegion]);

  return (
    <div className={styles.card}>
      <div className={styles.toolbar}>
        <h3 className={styles.title}>Employees ({sorted.length})</h3>
        <RegionFilter regions={regions} selectedRegion={selectedRegion} onChange={setSelectedRegion} />
      </div>

      {sorted.length === 0 ? (
        <p className={styles.empty}>No employees loaded yet.</p>
      ) : (
        <div className={styles.tableWrap}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th>Name</th>
                <th>Job Title</th>
                <th>Region</th>
              </tr>
            </thead>
            <tbody>
              {sorted.map(e => (
                <tr key={e.id}>
                  <td>{e.name}</td>
                  <td>{e.jobTitle}</td>
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
