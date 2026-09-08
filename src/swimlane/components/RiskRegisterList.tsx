import * as React from 'react';
import { DefaultButton, SearchBox } from '@fluentui/react';
import { IRiskStatement } from '../models/IRiskStatement';
import styles from './RiskRegisterList.module.scss';

export interface IRiskRegisterListProps {
  riskStatements: IRiskStatement[];
  onAddClick: () => void;
}

// Free-text filter across every column rather than per-column dropdowns -
// simplest thing that helps once the register has enough rows that
// scrolling to find one gets tedious (see user request).
function matchesQuery(risk: IRiskStatement, query: string): boolean {
  const q = query.trim().toLowerCase();
  if (!q) return true;
  return [
    risk.riskId, risk.category, risk.apqcProcessArea, risk.process, risk.function,
    risk.riskStatement, risk.rootCause, risk.riskResponse, risk.riskOwner
  ].some(field => field.toLowerCase().includes(q));
}

const RiskRegisterList: React.FC<IRiskRegisterListProps> = ({ riskStatements, onAddClick }) => {
  const [query, setQuery] = React.useState('');

  if (riskStatements.length === 0) {
    return (
      <div className={styles.card}>
        <div className={styles.toolbar}>
          <h3 className={styles.title}>Risk Register</h3>
          <DefaultButton text="+ Add risk" iconProps={{ iconName: 'Warning' }} onClick={onAddClick} />
        </div>
        <p className={styles.muted}>No risks loaded yet.</p>
      </div>
    );
  }

  const filtered = riskStatements.filter(r => matchesQuery(r, query));

  return (
    <div className={styles.card}>
      <div className={styles.toolbar}>
        <h3 className={styles.title}>Risk Register ({filtered.length} of {riskStatements.length})</h3>
        <div className={styles.toolbarActions}>
          <SearchBox
            placeholder="Search all columns..."
            value={query}
            onChange={(_e, v) => setQuery(v || '')}
            onClear={() => setQuery('')}
            className={styles.searchBox}
          />
          <DefaultButton text="+ Add risk" iconProps={{ iconName: 'Warning' }} onClick={onAddClick} />
        </div>
      </div>
      {filtered.length === 0 ? (
        <p className={styles.muted}>No risks match "{query}".</p>
      ) : (
        <div className={styles.tableWrap}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th>Risk ID</th>
                <th>Category</th>
                <th>APQC Process Area</th>
                <th>Process</th>
                <th>Function</th>
                <th>Risk Statement</th>
                <th>Root Cause</th>
                <th>Risk Response</th>
                <th>Risk Owner</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map(r => (
                <tr key={r.id}>
                  <td>{r.riskId || <span className={styles.muted}>—</span>}</td>
                  <td>{r.category || <span className={styles.muted}>—</span>}</td>
                  <td>{r.apqcProcessArea || <span className={styles.muted}>—</span>}</td>
                  <td>{r.process || <span className={styles.muted}>—</span>}</td>
                  <td>{r.function || <span className={styles.muted}>—</span>}</td>
                  <td>{r.riskStatement}</td>
                  <td className={styles.muted}>{r.rootCause || '—'}</td>
                  <td>{r.riskResponse || <span className={styles.muted}>—</span>}</td>
                  <td>{r.riskOwner || <span className={styles.muted}>—</span>}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
};

export default RiskRegisterList;
