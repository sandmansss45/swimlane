import * as React from 'react';
import { TextField, Dropdown, IDropdownOption, DefaultButton, IconButton } from '@fluentui/react';
import { IProcessStep, getShapeType } from '../models/IProcessStep';
import { IEmployee } from '../models/IEmployee';
import { IResolvedEdge } from '../utils/dependencyResolution';
import { buildColumns, columnKeyFor, compareProcessStepIds } from '../utils/columns';
import { connectorPath, gutterPath, pathMidpoint, rectFromDomRect, IRect } from '../utils/arrowRouting';
import ShapeNode from './shapes/ShapeNode';
import EmployeePicker from './EmployeePicker';
import styles from './SwimlaneCanvas.module.scss';

const SHAPE_OPTIONS: IDropdownOption[] = [
  { key: '', text: '(use Action Type / wording heuristic)' },
  { key: 'Process Step', text: 'Process (rounded rectangle)' },
  { key: 'Decision', text: 'Decision (diamond)' },
  { key: 'Approval', text: 'Approval (circle)' }
];

export interface ISwimlaneCanvasProps {
  steps: IProcessStep[]; // already filtered to the current Progress ID (and Process Step ID, if drilled down) - for display
  edges: IResolvedEdge[]; // resolved against the FULL, unfiltered dataset (row numbers only make sense that way) - this component only draws the ones whose endpoints are currently rendered
  drilledDownStepId: string | undefined;
  employees: IEmployee[];
  selectedRegion: string | undefined;
  onLabelEdge: (toRowId: string, token: string, label: string) => void;
  onEditStep: (step: IProcessStep) => void;
  onDeleteStep: (stepId: string) => void;
}

interface IEdgeGeometry {
  edge: IResolvedEdge;
  path: string;
  labelX: number;
  labelY: number;
}

// Ported from the earlier proven vanilla-JS renderer: lanes =
// ResponsibleJobTitle, columns = Process Step ID (or per-row when drilled
// into one step), shapes per the confirmed rules, arrows measured via DOM
// refs after layout and routed via connectorPath (picks whichever side of
// each box actually faces the other node, so lines don't cut diagonally
// through unrelated boxes between them).
interface IEditDraft {
  actionDescription: string;
  actionType: string;
  shapeOverride: string;
  responsibleJobTitle: string;
}

const SwimlaneCanvas: React.FC<ISwimlaneCanvasProps> = ({
  steps, edges, drilledDownStepId, employees, selectedRegion, onLabelEdge, onEditStep, onDeleteStep
}) => {
  const canvasRef = React.useRef<HTMLDivElement>(null);
  const nodeRefs = React.useRef(new Map<string, HTMLDivElement>());
  const [edgeGeometry, setEdgeGeometry] = React.useState<IEdgeGeometry[]>([]);
  const [selectedNodeId, setSelectedNodeId] = React.useState<string | undefined>(undefined);
  const [draftLabels, setDraftLabels] = React.useState<{ [token: string]: string }>({});
  const [editDraft, setEditDraft] = React.useState<IEditDraft | undefined>(undefined);

  const stepsById = React.useMemo(() => new Map(steps.map(s => [s.id, s])), [steps]);

  const lanes = React.useMemo(
    () => Array.from(new Set(steps.map(s => s.responsibleJobTitle || 'Unassigned'))),
    [steps]
  );
  const columns = React.useMemo(() => buildColumns(steps, drilledDownStepId), [steps, drilledDownStepId]);

  // Only edges whose both ends are currently rendered - the rest belong to
  // a different Progress ID / Process Step ID that isn't in view right now.
  const visibleEdges = React.useMemo(
    () => edges.filter(e => stepsById.has(e.fromRowId) && stepsById.has(e.toRowId)),
    [edges, stepsById]
  );

  const connectionsForSelectedNode = React.useMemo(
    () => selectedNodeId ? visibleEdges.filter(e => e.fromRowId === selectedNodeId) : [],
    [visibleEdges, selectedNodeId]
  );

  const measureEdges = React.useCallback(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const canvasRect = canvas.getBoundingClientRect();
    const geometries: IEdgeGeometry[] = [];
    const GUTTER_X = 30; // inside the 170px lane-label column, left of its text

    visibleEdges.forEach(edge => {
      const fromEl = nodeRefs.current.get(edge.fromRowId);
      const toEl = nodeRefs.current.get(edge.toRowId);
      const fromStep = stepsById.get(edge.fromRowId);
      const toStep = stepsById.get(edge.toRowId);
      if (!fromEl || !toEl || !fromStep || !toStep) return;

      const a: IRect = rectFromDomRect(fromEl.getBoundingClientRect(), canvasRect, canvas.scrollLeft, canvas.scrollTop);
      const b: IRect = rectFromDomRect(toEl.getBoundingClientRect(), canvasRect, canvas.scrollLeft, canvas.scrollTop);
      const mid = pathMidpoint(a, b);

      // Lanes stack vertically within a column, so a same-column
      // connection between different lanes almost always has at least one
      // other lane's box physically between source and target - route
      // those through the empty gutter instead of straight through
      // whatever's in the way. Same-lane (adjacent, nothing between them)
      // and different-column connections use the normal direct routing.
      const sameColumn = columnKeyFor(fromStep, drilledDownStepId) === columnKeyFor(toStep, drilledDownStepId);
      const sameLane = (fromStep.responsibleJobTitle || 'Unassigned') === (toStep.responsibleJobTitle || 'Unassigned');
      const path = (sameColumn && !sameLane) ? gutterPath(a, b, GUTTER_X) : connectorPath(a, b);

      geometries.push({ edge, path, labelX: mid.x, labelY: mid.y });
    });
    setEdgeGeometry(geometries);
  }, [visibleEdges, stepsById, drilledDownStepId]);

  React.useLayoutEffect(() => {
    measureEdges();
    window.addEventListener('resize', measureEdges);
    return () => window.removeEventListener('resize', measureEdges);
  }, [measureEdges]);

  const handleNodeClick = (step: IProcessStep): void => {
    const deselecting = selectedNodeId === step.id;
    setSelectedNodeId(deselecting ? undefined : step.id);
    setDraftLabels({});
    setEditDraft(deselecting ? undefined : {
      actionDescription: step.actionDescription,
      actionType: step.actionType,
      shapeOverride: step.shapeOverride || '',
      responsibleJobTitle: step.responsibleJobTitle
    });
  };

  const saveLabel = (token: string, toRowId: string): void => {
    const label = draftLabels[token];
    if (label && label.trim()) {
      onLabelEdge(toRowId, token, label.trim());
    }
  };

  const saveEdit = (): void => {
    const original = selectedNodeId ? stepsById.get(selectedNodeId) : undefined;
    if (!original || !editDraft) return;
    onEditStep({ ...original, ...editDraft });
  };

  const deleteSelected = (): void => {
    if (!selectedNodeId) return;
    onDeleteStep(selectedNodeId);
    setSelectedNodeId(undefined);
    setEditDraft(undefined);
  };

  if (steps.length === 0) {
    return <div className={styles.emptyState}>No steps to show yet.</div>;
  }

  return (
    <div className={styles.canvas} ref={canvasRef}>
      <svg className={styles.edgeOverlay}>
        <defs>
          <marker id="swimlaneArrowhead" markerWidth="9" markerHeight="9" refX="7" refY="4.5" orient="auto">
            <path d="M0,0 L9,4.5 L0,9 Z" fill="#000" />
          </marker>
          <marker id="swimlaneArrowheadHighlighted" markerWidth="9" markerHeight="9" refX="7" refY="4.5" orient="auto">
            <path d="M0,0 L9,4.5 L0,9 Z" fill="#ffb400" />
          </marker>
        </defs>
        {edgeGeometry.map(({ edge, path, labelX, labelY }) => {
          const highlighted = selectedNodeId !== undefined && edge.fromRowId === selectedNodeId;
          return (
            <g key={`${edge.fromRowId}-${edge.toRowId}-${edge.token}`}>
              <path
                d={path}
                stroke={highlighted ? '#ffb400' : '#000'}
                strokeWidth={highlighted ? 2.5 : 1.5}
                fill="none"
                markerEnd={highlighted ? 'url(#swimlaneArrowheadHighlighted)' : 'url(#swimlaneArrowhead)'}
              />
              {edge.label && (
                <text x={labelX} y={labelY - 4} textAnchor="middle" className={styles.edgeLabel}>{edge.label}</text>
              )}
            </g>
          );
        })}
      </svg>

      <div className={styles.grid} style={{ gridTemplateColumns: `170px repeat(${columns.length}, minmax(160px, 1fr))` }}>
        <div className={styles.corner} />
        {columns.map(col => (
          <div className={styles.columnHeader} key={col}>
            {drilledDownStepId ? (steps.find(s => s.processStepId === col)?.processStepName || col) : col}
          </div>
        ))}

        {lanes.map(lane => (
          <React.Fragment key={lane}>
            <div className={styles.laneLabel}>{lane}</div>
            {columns.map(col => {
              const cellSteps = steps
                .filter(s => (s.responsibleJobTitle || 'Unassigned') === lane && columnKeyFor(s, drilledDownStepId) === col)
                .sort((a, b) => compareProcessStepIds(a.processStepId, b.processStepId));
              return (
                <div className={styles.laneCell} key={`${lane}-${col}`}>
                  {cellSteps.map(step => (
                    <div
                      key={step.id}
                      ref={el => { if (el) nodeRefs.current.set(step.id, el); }}
                    >
                      <ShapeNode
                        label={step.actionDescription}
                        shape={getShapeType(step)}
                        selected={selectedNodeId === step.id}
                        onClick={() => handleNodeClick(step)}
                      />
                    </div>
                  ))}
                </div>
              );
            })}
          </React.Fragment>
        ))}
      </div>

      {selectedNodeId && editDraft && (
        <div className={styles.connectionsPanel}>
          <div className={styles.editHeader}>
            <h4>Edit task</h4>
            <IconButton iconProps={{ iconName: 'Delete' }} title="Delete this task" onClick={deleteSelected} />
          </div>

          <div className={styles.editForm}>
            <TextField
              label="Action description"
              multiline
              value={editDraft.actionDescription}
              onChange={(_e, v) => setEditDraft(prev => (prev ? { ...prev, actionDescription: v || '' } : prev))}
            />
            <TextField
              label="Action type"
              value={editDraft.actionType}
              onChange={(_e, v) => setEditDraft(prev => (prev ? { ...prev, actionType: v || '' } : prev))}
            />
            <Dropdown
              label="Shape"
              selectedKey={editDraft.shapeOverride}
              options={SHAPE_OPTIONS}
              onChange={(_e, option) => setEditDraft(prev => (prev && option ? { ...prev, shapeOverride: String(option.key) } : prev))}
            />
            <EmployeePicker
              employees={employees}
              selectedRegion={selectedRegion}
              actionType={editDraft.actionType}
              value={editDraft.responsibleJobTitle}
              onChange={(jobTitle) => setEditDraft(prev => (prev ? { ...prev, responsibleJobTitle: jobTitle } : prev))}
            />
            <DefaultButton text="Save changes" onClick={saveEdit} />
          </div>

          <h4>Outgoing connections</h4>
          {connectionsForSelectedNode.length === 0 && <p>Nothing else currently visible depends on this row.</p>}
          {connectionsForSelectedNode.map(edge => (
            <div className={styles.connectionRow} key={edge.token}>
              <span>&rarr; {stepsById.get(edge.toRowId)?.actionDescription}</span>
              <input
                type="text"
                placeholder="Yes / No / label this branch"
                defaultValue={edge.label || ''}
                onChange={(e) => setDraftLabels(prev => ({ ...prev, [edge.token]: e.target.value }))}
                onBlur={() => saveLabel(edge.token, edge.toRowId)}
              />
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

export default SwimlaneCanvas;
