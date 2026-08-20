import * as React from 'react';
import { SwimlaneStage } from '../models/ISwimlaneStatus';
import styles from './ShapeLegend.module.scss';

export interface IShapeLegendProps {
  // Drag-to-create: dragging a "Step type" swatch and dropping it onto a
  // cell in the canvas grid creates a new step there with that shape
  // already set - see SwimlaneCanvas's draggingNewShape/onCreateStep.
  // Values match ProcessStepForm's SHAPE_OPTIONS keys exactly ('Process
  // Step' | 'Decision' | 'Approval' | 'Document'), since they're written
  // straight into the new step's shapeOverride field.
  onDragShapeStart: (shapeOverride: string) => void;
  onDragShapeEnd: () => void;
  // Draft/Finalised status for the swimlane on screen right now (see
  // ISwimlaneStatus) - shown underneath the rest of the legend with a
  // toggle button. Purely a display label; never gates anything.
  stage: SwimlaneStage;
  // Undefined means no one has ever explicitly set a stage - Draft shows
  // with no attribution rather than a misleading "set by nobody" line.
  stageSetBy: string | undefined;
  onToggleStage: () => void;
}

// Step type is conveyed by both FORM (mirrors getShapeType in
// models/IProcessStep.ts) and, for Decision/Approval, FILL COLOR too -
// both plain, unconditional lists rather than derived from the data,
// since a legend should show what every possible shape/color means
// regardless of which ones happen to appear in the steps on screen.
const ShapeLegend: React.FC<IShapeLegendProps> = ({ onDragShapeStart, onDragShapeEnd, stage, stageSetBy, onToggleStage }) => {
  const dragProps = (shapeOverride: string): React.HTMLAttributes<HTMLSpanElement> => ({
    draggable: true,
    onDragStart: e => {
      e.dataTransfer.effectAllowed = 'copy';
      // Firefox refuses to start a drag at all unless setData is called -
      // same reasoning as SwimlaneCanvas's own handleDragStart.
      e.dataTransfer.setData('text/plain', shapeOverride);
      onDragShapeStart(shapeOverride);
    },
    onDragEnd: onDragShapeEnd
  });

  return (
    <div className={styles.legend}>
      <div className={styles.row}>
        <div className={styles.group}>
          <span className={styles.groupLabel}>Step type - drag onto the flow to add</span>
          <span className={`${styles.item} ${styles.draggableItem}`} title="Drag onto the flow to add a Process step" {...dragProps('Process Step')}>
            <span className={`${styles.swatch} ${styles.process}`} />
            Process step
          </span>
          <span className={`${styles.item} ${styles.draggableItem}`} title="Drag onto the flow to add a Decision" {...dragProps('Decision')}>
            <span className={`${styles.swatch} ${styles.decision}`} />
            Decision
          </span>
          <span className={`${styles.item} ${styles.draggableItem}`} title="Drag onto the flow to add an Approval" {...dragProps('Approval')}>
            <span className={`${styles.swatch} ${styles.approval}`} />
            Approval
          </span>
          <span className={`${styles.item} ${styles.draggableItem}`} title="Drag onto the flow to add a Document" {...dragProps('Document')}>
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

      {/*
        Its own row underneath the rest of the key, per explicit user
        request - not just another inline group in the row above, since
        that would read as one more shape-legend entry rather than a
        distinct status for the whole swimlane.
      */}
      <div className={styles.stageRow}>
        <span className={`${styles.stageBadge} ${stage === 'Finalised' ? styles.stageFinalised : styles.stageDraft}`}>
          {stage}
        </span>
        {stageSetBy && <span className={styles.stageBy}>set by {stageSetBy}</span>}
        <button type="button" className={styles.stageToggle} onClick={onToggleStage}>
          Mark as {stage === 'Draft' ? 'Finalised' : 'Draft'}
        </button>
      </div>
    </div>
  );
};

export default ShapeLegend;
