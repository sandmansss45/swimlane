import * as React from 'react';
import { ShapeType } from '../../models/IProcessStep';
import { RiskLevel } from '../../models/IRiskStatement';
import styles from './ShapeNode.module.scss';

export interface IShapeNodeProps {
  label: string;
  shape: ShapeType;
  riskLevel?: RiskLevel; // when set, overrides the default fill with the traffic-light risk color (unrelated ad hoc flag, see riskLevelOverride)
  // Worst severity among this step's real Risk Register links, if any -
  // renders as a small corner marker rather than recoloring the whole
  // shape, so it never collides visually with riskLevel above.
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

// Shape TYPE is conveyed by form only: approval = circles, decision =
// diamonds, ordinary process steps = rounded rectangles, document-type
// (a physical/system artifact like a PO or invoice passing between
// people, not an action someone takes) = wavy-bottomed rectangles.
// Confirmed design rule - do not flip this. Fill COLOR is separate: blue
// by default, or red/amber/green once riskLevelOverride is set - see
// ShapeNode.module.scss.
const ShapeNode: React.FC<IShapeNodeProps> = ({ label, shape, riskLevel, linkedRiskSeverity, linkedRiskCount, selected, onClick }) => {
  const riskClass = riskLevel === 'High' ? styles.riskHigh
    : riskLevel === 'Medium' ? styles.riskMedium
    : riskLevel === 'Low' ? styles.riskLow
    : '';

  const className = [
    styles.shapeNode,
    shape === 'decision' ? styles.decision : '',
    shape === 'approval' ? styles.approval : '',
    shape === 'process' ? styles.process : '',
    shape === 'document' ? styles.document : '',
    riskClass,
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
