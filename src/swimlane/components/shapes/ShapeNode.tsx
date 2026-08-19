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
  // Drag-to-connect: a small handle rendered only when this is provided,
  // dragging it starts a new dependency link FROM this step - see
  // connectingFromStepId in SwimlaneCanvas.tsx. Kept as its own nested
  // draggable element rather than reusing the node's own body-drag (which
  // already means "move this step to a different lane/position") so the
  // two gestures can never collide - the browser resolves a drag started
  // exactly on the handle to the handle's own dragstart, not the parent's.
  onConnectorDragStart?: () => void;
  // Fires on release regardless of where the drop landed (or if it landed
  // on nothing valid at all) - the one place connection-drag state
  // reliably gets cleared, same role handleDragEnd plays for the
  // existing move-drag.
  onConnectorDragEnd?: () => void;
  // True while a connection drag is in progress AND this specific step is
  // a valid place to drop it (see isValidConnectionTarget) - highlighted
  // the same way a valid drop cell is elsewhere, so it's clear up front
  // where dropping is even possible before the pointer gets there.
  connectable?: boolean;
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
const ShapeNode: React.FC<IShapeNodeProps> = ({
  label, shape, linkedRiskSeverity, linkedRiskCount, selected, onClick, onConnectorDragStart, onConnectorDragEnd, connectable
}) => {
  const className = [
    styles.shapeNode,
    shape === 'decision' ? styles.decision : '',
    shape === 'approval' ? styles.approval : '',
    shape === 'process' ? styles.process : '',
    shape === 'document' ? styles.document : '',
    selected ? styles.selected : '',
    connectable ? styles.connectable : ''
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
      {onConnectorDragStart && (
        <span
          className={styles.connectorHandle}
          draggable
          title="Drag to another step to link them"
          // Native drag needs the drag itself to actually start from THIS
          // element, not just a click - onDragStart is what fires that
          // moment; the click handler above is deliberately left alone so
          // a plain click here still opens the edit panel like anywhere
          // else on the shape (dragging and clicking are different
          // gestures on the same pixel, exactly like the node body's own
          // move-drag vs. its click-to-edit already work).
          onDragStart={e => { e.stopPropagation(); onConnectorDragStart(); }}
          onDragEnd={e => { e.stopPropagation(); onConnectorDragEnd?.(); }}
          onClick={e => e.stopPropagation()}
        >
          →
        </span>
      )}
    </div>
  );
};

export default ShapeNode;
