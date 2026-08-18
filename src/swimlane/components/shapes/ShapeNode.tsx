import * as React from 'react';
import { ShapeType } from '../../models/IProcessStep';
import { RiskLevel } from '../../models/IRiskStatement';
import styles from './ShapeNode.module.scss';

export interface IShapeNodeProps {
  label: string;
  shape: ShapeType;
  // Worst severity among this step's real Risk Register links, if any -
  // renders as a small corner marker, the app's one real risk indicator
  // (a separate manual "risk level" fill override used to exist here too,
  // but was retired - see ShapeNode.module.scss).
  linkedRiskSeverity?: RiskLevel;
  linkedRiskCount?: number;
  selected?: boolean; // true while this node's connections are being shown/labeled
  onClick?: () => void;
}

const MARKER_CLASS: Record<RiskLevel, string> = {
  High: styles.markerHigh,
  Medium: styles.markerMedium,
  Low: styles.markerLow
};

// Shape TYPE is conveyed by form (approval = circles, decision = diamonds,
// ordinary process steps = rounded rectangles, document-type = wavy-
// bottomed rectangles) AND, for Decision/Approval specifically, by fill
// color too (orange/green) so they stand out at a glance among a diagram
// full of default-blue process/document shapes - see ShapeNode.module.scss.
const ShapeNode: React.FC<IShapeNodeProps> = ({ label, shape, linkedRiskSeverity, linkedRiskCount, selected, onClick }) => {
  const className = [
    styles.shapeNode,
    shape === 'decision' ? styles.decision : '',
    shape === 'approval' ? styles.approval : '',
    shape === 'process' ? styles.process : '',
    shape === 'document' ? styles.document : '',
    selected ? styles.selected : ''
  ].filter(Boolean).join(' ');

  return (
    <div className={className} onClick={onClick} title={label}>
      <span className={styles.label}>{label}</span>
      {linkedRiskSeverity && (
        <span
          className={`${styles.riskMarker} ${MARKER_CLASS[linkedRiskSeverity]}`}
          title={`${linkedRiskCount} linked risk${linkedRiskCount === 1 ? '' : 's'} (worst: ${linkedRiskSeverity})`}
        >
          !
        </span>
      )}
    </div>
  );
};

export default ShapeNode;
