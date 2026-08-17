import * as React from 'react';
import { IRiskStatement } from '../models/IRiskStatement';
import { IProcessStep } from '../models/IProcessStep';
import styles from './RiskRegisterList.module.scss';

export interface IRiskRegisterListProps {
  riskStatements: IRiskStatement[];
  // Used to work out which steps currently link each risk - the link
  // itself lives on the step (see IProcessStep.linkedRisks), not here, so
  // this table has to look it up rather than just reading a field.
  steps: IProcessStep[];
}

const RiskRegisterList: React.FC<IRiskRegisterListProps> = ({ riskStatements, steps }) => {
  const linkedStepIdsByRisk = React.useMemo(() => {
    const map = new Map<string, string[]>();
    steps.forEach(step => {
      (step.linkedRisks || []).forEach(link => {
        const list = map.get(link.riskId) || [];
        list.push(step.processStepId);
        map.set(link.riskId, list);
      });
    });
    return map;
  }, [steps]);

  if (riskStatements.length === 0) {
    return (
      <div className={styles.card}>
        <h3 className={styles.title}>Risk Register</h3>
        <p className={styles.muted}>No risks loaded yet.</p>
      </div>
    );
  }

  return (
    <div className={styles.card}>
      <h3 className={styles.title}>Risk Register ({riskStatements.length})</h3>
      <div className={styles.tableWrap}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th>Risk ID</th>
              <th>Category</th>
              <th>Risk Statement</th>
              <th>Root Cause</th>
              <th>Likelihood</th>
              <th>Materiality ($Mn)</th>
              <th>Inherent Risk Rating</th>
              <th>Risk Response</th>
              <th>Linked steps</th>
            </tr>
          </thead>
          <tbody>
            {riskStatements.map(r => {
              const linkedStepIds = linkedStepIdsByRisk.get(r.id) || [];
              return (
                <tr key={r.id}>
                  <td>{r.riskId || <span className={styles.muted}>—</span>}</td>
                  <td>{r.category || <span className={styles.muted}>—</span>}</td>
                  <td>{r.riskStatement}</td>
                  <td className={styles.muted}>{r.rootCause || '—'}</td>
                  <td>{r.likelihood !== undefined ? r.likelihood : <span className={styles.muted}>—</span>}</td>
                  <td>{r.materiality !== undefined ? r.materiality : <span className={styles.muted}>—</span>}</td>
                  <td>{r.inherentRiskRating !== undefined ? r.inherentRiskRating : <span className={styles.muted}>—</span>}</td>
                  <td>{r.riskResponse || <span className={styles.muted}>—</span>}</td>
                  <td>{linkedStepIds.length > 0 ? linkedStepIds.join(', ') : <span className={styles.muted}>—</span>}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default RiskRegisterList;
