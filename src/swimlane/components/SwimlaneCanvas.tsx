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
  { key: 'Approval', text: 'Approval (circle)' },
  { key: 'Document', text: 'Document (artifact)' }
];

const truncate = (text: string, max: number): string =>
  text.length > max ? `${text.slice(0, max - 1)}…` : text;

export interface ISwimlaneCanvasProps {
  steps: IProcessStep[]; // already filtered to the current Progress ID (and Process Step ID, if drilled down) - for display
  allSteps: IProcessStep[]; // FULL, unfiltered, original-order dataset - needed to compute row-number DependsOn tokens and to offer every step as a "depends on" option regardless of what's currently filtered into view
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
  dependsOnStepIds: string[];
}

const SwimlaneCanvas: React.FC<ISwimlaneCanvasProps> = ({
  steps, allSteps, edges, drilledDownStepId, employees, selectedRegion, onLabelEdge, onEditStep, onDeleteStep
}) => {
  const canvasRef = React.useRef<HTMLDivElement>(null);
  const nodeRefs = React.useRef(new Map<string, HTMLDivElement>());
  const [edgeGeometry, setEdgeGeometry] = React.useState<IEdgeGeometry[]>([]);
  const [selectedNodeId, setSelectedNodeId] = React.useState<string | undefined>(undefined);
  const [draftLabels, setDraftLabels] = React.useState<{ [token: string]: string }>({});
  const [editDraft, setEditDraft] = React.useState<IEditDraft | undefined>(undefined);

  const stepsById = React.useMemo(() => new Map(steps.map(s => [s.id, s])), [steps]);

  // Row number = position in the FULL unfiltered dataset + 2 (header
  // counted as row 1) - the same scheme resolveDependencyEdges uses.
  // Needed here so the "Depends on" picker can convert a human pick
  // (another step) into the row-number token the rest of the app expects,
  // and back again when loading a step's existing dependencies into the form.
  const rowNumberById = React.useMemo(() => {
    const map = new Map<string, number>();
    allSteps.forEach((s, i) => map.set(s.id, i + 2));
    return map;
  }, [allSteps]);

  const stepIdByRowNumber = React.useMemo(() => {
    const map = new Map<number, string>();
    allSteps.forEach((s, i) => map.set(i + 2, s.id));
    return map;
  }, [allSteps]);

  const dependsOnOptions: IDropdownOption[] = React.useMemo(
    () => allSteps
      .filter(s => s.id !== selectedNodeId)
      .map(s => ({ key: s.id, text: `${s.processStepId} — ${truncate(s.actionDescription, 50)}` })),
    [allSteps, selectedNodeId]
  );

  const lanes = React.useMemo(
    () => Array.from(new Set(steps.map(s => s.responsibleJobTitle || 'Unassigned'))),
    [steps]
  );
  const columns = React.useMemo(() => buildColumns(steps, drilledDownStepId), [steps, drilledDownStepId]);

  // Each step's position within its own lane+column cell's stacking order
  // (same sort the render loop below uses) - needed to tell whether two
  // steps in the SAME lane are adjacent (nothing between them, a direct
  // line is fine) or not (another step from that same lane sits between
  // them in the stack, so a direct line would cut straight through it -
  // this is the same problem gutter-routing already solves for
  // different-lane connections, just triggered by lane+column stacking
  // order instead of by lane identity).
  const cellStackIndex = React.useMemo(() => {
    const index = new Map<string, number>();
    lanes.forEach(lane => {
      columns.forEach(col => {
        const cellSteps = steps
          .filter(s => (s.responsibleJobTitle || 'Unassigned') === lane && columnKeyFor(s, drilledDownStepId) === col)
          .sort((a, b) => compareProcessStepIds(a.processStepId, b.processStepId));
        cellSteps.forEach((s, i) => index.set(s.id, i));
      });
    });
    return index;
  }, [steps, lanes, columns, drilledDownStepId]);

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

      // A direct line is only safe between two boxes that are truly
      // adjacent in the SAME lane+column cell's stack - literally nothing
      // else can be physically between them. Any other pairing (different
      // lane, different column, or same lane/column but with another step
      // stacked between them) risks the connector's bend cutting straight
      // through unrelated boxes sitting on that path - route those
      // through the empty left-margin gutter instead, which is guaranteed
      // clear no matter how far apart the two boxes are.
      const sameColumn = columnKeyFor(fromStep, drilledDownStepId) === columnKeyFor(toStep, drilledDownStepId);
      const sameLane = (fromStep.responsibleJobTitle || 'Unassigned') === (toStep.responsibleJobTitle || 'Unassigned');
      const fromIdx = cellStackIndex.get(fromStep.id);
      const toIdx = cellStackIndex.get(toStep.id);
      const adjacentInSameCell = sameLane && sameColumn && fromIdx !== undefined && toIdx !== undefined && Math.abs(fromIdx - toIdx) <= 1;
      const needsGutter = !adjacentInSameCell;
      const path = needsGutter ? gutterPath(a, b, GUTTER_X) : connectorPath(a, b);

      geometries.push({ edge, path, labelX: mid.x, labelY: mid.y });
    });
    setEdgeGeometry(geometries);
  }, [visibleEdges, stepsById, drilledDownStepId, cellStackIndex]);

  React.useLayoutEffect(() => {
    measureEdges();
    window.addEventListener('resize', measureEdges);
    return () => window.removeEventListener('resize', measureEdges);
  }, [measureEdges]);

  const resolveDependsOnStepIds = (step: IProcessStep): string[] => {
    const ids: string[] = [];
    step.dependsOn.forEach(token => {
      const match = /(\d+)\s*$/.exec(token);
      if (!match) return;
      const depId = stepIdByRowNumber.get(parseInt(match[1], 10));
      if (depId) ids.push(depId);
    });
    return ids;
  };

  const handleNodeClick = (step: IProcessStep): void => {
    const deselecting = selectedNodeId === step.id;
    setSelectedNodeId(deselecting ? undefined : step.id);
    setDraftLabels({});
    setEditDraft(deselecting ? undefined : {
      actionDescription: step.actionDescription,
      actionType: step.actionType,
      shapeOverride: step.shapeOverride || '',
      responsibleJobTitle: step.responsibleJobTitle,
      dependsOnStepIds: resolveDependsOnStepIds(step)
    });
  };

  // Two different dependent steps can produce the exact same literal
  // DependsOn token text (e.g. both depending on the same decision row
  // while sharing that decision's own processStepId prefix), so the
  // token alone isn't a safe key for draft input state or React lists -
  // scope every lookup to toRowId+token instead.
  const draftKey = (edge: IResolvedEdge): string => `${edge.toRowId}::${edge.token}`;

  const saveLabel = (edge: IResolvedEdge): void => {
    const label = draftLabels[draftKey(edge)];
    if (label && label.trim()) {
      onLabelEdge(edge.toRowId, edge.token, label.trim());
    }
  };

  const saveEdit = (): void => {
    const original = selectedNodeId ? stepsById.get(selectedNodeId) : undefined;
    if (!original || !editDraft) return;
    const { dependsOnStepIds, ...fields } = editDraft;
    const dependsOn = dependsOnStepIds
      .map(id => rowNumberById.get(id))
      .filter((rowNumber): rowNumber is number => rowNumber !== undefined)
      .map(rowNumber => String(rowNumber));
    onEditStep({ ...original, ...fields, dependsOn });
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
            <path d="M0,0 L9,4.5 L0,9 Z" fill="#3c4a63" />
          </marker>
          <marker id="swimlaneArrowheadHighlighted" markerWidth="9" markerHeight="9" refX="7" refY="4.5" orient="auto">
            <path d="M0,0 L9,4.5 L0,9 Z" fill="#1f4fa3" />
          </marker>
        </defs>
        {edgeGeometry.map(({ edge, path, labelX, labelY }) => {
          const highlighted = selectedNodeId !== undefined && edge.fromRowId === selectedNodeId;
          return (
            <g key={`${edge.fromRowId}-${edge.toRowId}-${edge.token}`}>
              <path
                d={path}
                stroke={highlighted ? '#1f4fa3' : '#3c4a63'}
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
            <Dropdown
              label="Depends on"
              placeholder="Which step(s) does this follow?"
              multiSelect
              selectedKeys={editDraft.dependsOnStepIds}
              options={dependsOnOptions}
              onChange={(_e, option) => setEditDraft(prev => {
                if (!prev || !option) return prev;
                const ids = option.selected
                  ? [...prev.dependsOnStepIds, String(option.key)]
                  : prev.dependsOnStepIds.filter(id => id !== option.key);
                return { ...prev, dependsOnStepIds: ids };
              })}
            />
            <DefaultButton text="Save changes" onClick={saveEdit} />
          </div>

          <h4>Outgoing connections</h4>
          {connectionsForSelectedNode.length === 0 && <p>Nothing else currently visible depends on this row.</p>}
          {connectionsForSelectedNode.map(edge => (
            <div className={styles.connectionRow} key={`${edge.toRowId}-${edge.token}`}>
              <span>&rarr; {stepsById.get(edge.toRowId)?.actionDescription}</span>
              <input
                type="text"
                placeholder="Yes / No / label this branch"
                defaultValue={edge.label || ''}
                onChange={(e) => setDraftLabels(prev => ({ ...prev, [draftKey(edge)]: e.target.value }))}
                onBlur={() => saveLabel(edge)}
              />
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

export default SwimlaneCanvas;
