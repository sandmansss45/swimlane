import * as React from 'react';
import { ShapeType } from '../../models/IProcessStep';
import { RiskLevel } from '../../models/IRiskStatement';
import styles from './ShapeNode.module.scss';

export interface IShapeNodeProps {
  label: string;
  shape: ShapeType;
  riskLevel?: RiskLevel; // when set, overrides the default fill with the traffic-light risk color
  selected?: boolean; // true while this node's connections are being shown/labeled
  onClick?: () => void;
}

// Shape TYPE is conveyed by form only: approval = circles, decision =
// diamonds, ordinary process steps = rounded rectangles, document-type
// (a physical/system artifact like a PO or invoice passing between
// people, not an action someone takes) = wavy-bottomed rectangles.
// Confirmed design rule - do not flip this. Fill COLOR is separate: blue
// by default, or red/amber/green once a Risk Register entry links to
// this step - see ShapeNode.module.scss.
const ShapeNode: React.FC<IShapeNodeProps> = ({ label, shape, riskLevel, selected, onClick }) => {
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
    </div>
  );
};

export default ShapeNode;
