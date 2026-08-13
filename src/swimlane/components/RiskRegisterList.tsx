import * as React from 'react';
import { IRiskStatement } from '../models/IRiskStatement';
import styles from './RiskRegisterList.module.scss';

export interface IRiskRegisterListProps {
  riskStatements: IRiskStatement[];
}

const LEVEL_CLASS: Record<string, string> = {
  High: styles.levelHigh,
  Medium: styles.levelMedium,
  Low: styles.levelLow
};

const RiskRegisterList: React.FC<IRiskRegisterListProps> = ({ riskStatements }) => {
  if (riskStatements.length === 0) return null;

  return (
    <div className={styles.card}>
      <h3 className={styles.title}>Risk Register ({riskStatements.length})</h3>
      <div className={styles.tableWrap}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th>Level</th>
              <th>Risk</th>
              <th>Statement</th>
              <th>Linked steps</th>
            </tr>
          </thead>
          <tbody>
            {riskStatements.map(r => (
              <tr key={r.id}>
                <td>
                  {r.riskLevel ? (
                    <span className={`${styles.badge} ${LEVEL_CLASS[r.riskLevel]}`}>{r.riskLevel}</span>
                  ) : (
                    <span className={styles.muted}>—</span>
                  )}
                </td>
                <td><strong>{r.title}</strong></td>
                <td>{r.riskStatement}</td>
                <td>{(r.linkedProcessStepIds || []).join(', ') || <span className={styles.muted}>—</span>}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default RiskRegisterList;
