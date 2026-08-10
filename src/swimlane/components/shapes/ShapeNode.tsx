import * as React from 'react';
import { ShapeType } from '../../models/IProcessStep';
import styles from './ShapeNode.module.scss';

export interface IShapeNodeProps {
  label: string;
  shape: ShapeType;
  selected?: boolean; // true while this node's connections are being shown/labeled
  onClick?: () => void;
}

// Approval-type = circles (green). Decision-type = diamonds (orange).
// Ordinary process steps = rounded rectangles (blue). Confirmed design
// rule - do not flip this.
const ShapeNode: React.FC<IShapeNodeProps> = ({ label, shape, selected, onClick }) => {
  const className = [
    styles.shapeNode,
    shape === 'decision' ? styles.decision : '',
    shape === 'approval' ? styles.approval : '',
    shape === 'process' ? styles.process : '',
    selected ? styles.selected : ''
  ].filter(Boolean).join(' ');

  return (
    <div className={className} onClick={onClick} title={label}>
      <span className={styles.label}>{label}</span>
    </div>
  );
};

export default ShapeNode;
