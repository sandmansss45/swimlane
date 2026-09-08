import * as React from 'react';
import { DefaultButton } from '@fluentui/react';
import { IRiskStatement } from '../models/IRiskStatement';
import styles from './RiskRegisterList.module.scss';

export interface IRiskRegisterListProps {
  riskStatements: IRiskStatement[];
  onAddClick: () => void;
}

const RiskRegisterList: React.FC<IRiskRegisterListProps> = ({ riskStatements, onAddClick }) => {
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

  return (
    <div className={styles.card}>
      <div className={styles.toolbar}>
        <h3 className={styles.title}>Risk Register ({riskStatements.length})</h3>
        <DefaultButton text="+ Add risk" iconProps={{ iconName: 'Warning' }} onClick={onAddClick} />
      </div>
      <div className={styles.tableWrap}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th>Risk ID</th>
              <th>Category</th>
              <th>Risk Statement</th>
              <th>Root Cause</th>
              <th>Risk Response</th>
              <th>Risk Owner</th>
            </tr>
          </thead>
          <tbody>
            {riskStatements.map(r => (
              <tr key={r.id}>
                <td>{r.riskId || <span className={styles.muted}>—</span>}</td>
                <td>{r.category || <span className={styles.muted}>—</span>}</td>
                <td>{r.riskStatement}</td>
                <td className={styles.muted}>{r.rootCause || '—'}</td>
                <td>{r.riskResponse || <span className={styles.muted}>—</span>}</td>
                <td>{r.riskOwner || <span className={styles.muted}>—</span>}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default RiskRegisterList;
