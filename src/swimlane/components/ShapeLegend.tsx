import * as React from 'react';
import styles from './ShapeLegend.module.scss';

// Step type is conveyed by both FORM (mirrors getShapeType in
// models/IProcessStep.ts) and, for Decision/Approval, FILL COLOR too -
// both plain, unconditional lists rather than derived from the data,
// since a legend should show what every possible shape/color means
// regardless of which ones happen to appear in the steps on screen.
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
      <span className={styles.groupLabel}>Linked risk</span>
      <span className={styles.item}>
        <span className={styles.markerSample}>!</span>
        Linked risk - click the step for details
      </span>
    </div>
  </div>
);

export default ShapeLegend;
