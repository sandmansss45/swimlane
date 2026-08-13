import * as React from 'react';
import styles from './ShapeLegend.module.scss';

// Two independent keys, since the two things a shape communicates are
// independent: FORM (mirrors getShapeType in models/IProcessStep.ts) says
// what kind of step it is, and FILL COLOR says its risk severity once the
// Risk Register links a risk to it (blue = no risk linked). Both are
// plain, unconditional lists rather than derived from the data, since a
// legend should show what every possible shape/color means regardless of
// which ones happen to appear in the steps currently on screen.
const ShapeLegend: React.FC = () => (
  <div className={styles.legend}>
    <div className={styles.group}>
      <span className={styles.groupLabel}>Step type</span>
      <span className={styles.item}>
        <span className={`${styles.swatch} ${styles.process}`} />
        Process step
      </span>
      <span className={styles.item}>
        <span className={`${styles.swatch} ${styles.decision}`} />
        Decision
      </span>
      <span className={styles.item}>
        <span className={`${styles.swatch} ${styles.approval}`} />
        Approval
      </span>
      <span className={styles.item}>
        <span className={`${styles.swatch} ${styles.document}`} />
        Document
      </span>
    </div>
    <div className={styles.divider} />
    <div className={styles.group}>
      <span className={styles.groupLabel}>Risk level</span>
      <span className={styles.item}>
        <span className={`${styles.dot} ${styles.riskHigh}`} />
        High
      </span>
      <span className={styles.item}>
        <span className={`${styles.dot} ${styles.riskMedium}`} />
        Medium
      </span>
      <span className={styles.item}>
        <span className={`${styles.dot} ${styles.riskLow}`} />
        Low
      </span>
    </div>
  </div>
);

export default ShapeLegend;
