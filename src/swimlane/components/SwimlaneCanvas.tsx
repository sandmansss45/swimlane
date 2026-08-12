import * as React from 'react';
import { TextField, Dropdown, IDropdownOption, DefaultButton, IconButton, Modal } from '@fluentui/react';
import { IProcessStep, getShapeType } from '../models/IProcessStep';
import { IEmployee } from '../models/IEmployee';
import { IResolvedEdge } from '../utils/dependencyResolution';
import { orderStepsForTimeline, buildColumnGroups } from '../utils/columns';
import { connectorPath, highwayPath, rectFromDomRect, IRect } from '../utils/arrowRouting';
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
  const svgRef = React.useRef<SVGSVGElement>(null);
  const highwaySpacerRef = React.useRef<HTMLDivElement>(null);
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

  // Every visible step gets its own column (timeline slot) instead of
  // being grouped/stacked with every other row that shares its Process
  // Step ID - see utils/columns.ts for why. orderedSteps IS the column
  // axis: orderedSteps[i] is column i.
  const orderedSteps = React.useMemo(() => orderStepsForTimeline(steps), [steps]);
  const columnGroups = React.useMemo(() => buildColumnGroups(orderedSteps), [orderedSteps]);

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
    const HIGHWAY_TRACK_Y = [0, -8, 8, -16, 16]; // parallel tracks inside the reserved highway strip

    // The overlay's width/height:100% would otherwise only cover the
    // CANVAS's visible viewport (its CSS containing-block size), not the
    // full scrollable grid - since path coordinates are computed in
    // full-content space (see rectFromDomRect's scroll offset add-back),
    // any edge past the initially-visible width/height would silently
    // fall outside the SVG's own box and get clipped by its default
    // overflow:hidden. Force it to the real scrollable size instead.
    if (svgRef.current) {
      svgRef.current.style.width = `${canvas.scrollWidth}px`;
      svgRef.current.style.height = `${canvas.scrollHeight}px`;
    }

    const rectById = new Map<string, IRect>();
    orderedSteps.forEach(step => {
      const el = nodeRefs.current.get(step.id);
      if (!el) return;
      rectById.set(step.id, rectFromDomRect(el.getBoundingClientRect(), canvasRect, canvas.scrollLeft, canvas.scrollTop));
    });
    const allRects = Array.from(rectById.entries());

    const highwayBaseY = highwaySpacerRef.current
      ? rectFromDomRect(highwaySpacerRef.current.getBoundingClientRect(), canvasRect, canvas.scrollLeft, canvas.scrollTop).cy
      : 20;

    // A direct line between two boxes is safe exactly when no OTHER box
    // sits anywhere inside the straight-line path's bounding rectangle -
    // connectorPath's bends never travel outside that rectangle, so this
    // is both a necessary and sufficient check, and it works regardless
    // of which lane/column the two boxes happen to land in (unlike a
    // "same lane" / "same column" heuristic, which only covers the grid
    // shapes this app happened to use before).
    const pathIsClear = (id1: string, a: IRect, id2: string, b: IRect): boolean => {
      const left = Math.min(a.left, b.left);
      const right = Math.max(a.right, b.right);
      const top = Math.min(a.top, b.top);
      const bottom = Math.max(a.bottom, b.bottom);
      return !allRects.some(([id, rect]) => {
        if (id === id1 || id === id2) return false;
        return rect.left < right && rect.right > left && rect.top < bottom && rect.bottom > top;
      });
    };

    interface ICandidate { edge: IResolvedEdge; a: IRect; b: IRect; needsHighway: boolean }
    const candidates: ICandidate[] = [];

    visibleEdges.forEach(edge => {
      const a = rectById.get(edge.fromRowId);
      const b = rectById.get(edge.toRowId);
      if (!a || !b) return;
      candidates.push({ edge, a, b, needsHighway: !pathIsClear(edge.fromRowId, a, edge.toRowId, b) });
    });

    // Every highway edge's horizontal run shares the same reserved strip,
    // so two edges whose X-ranges overlap would draw on top of each other
    // and read as one merged line - assign each a distinct parallel
    // Y-level within the strip (classic greedy interval-coloring: sorted
    // by start, reuse the first track that's already clear by then) so
    // overlapping connections stay visually distinguishable.
    const highwayEdges = candidates
      .filter(c => c.needsHighway)
      .map(c => ({ ...c, xStart: Math.min(c.a.cx, c.b.cx), xEnd: Math.max(c.a.cx, c.b.cx) }))
      .sort((x, y) => x.xStart - y.xStart);
    const trackEndX: number[] = [];
    const trackByEdge = new Map<IResolvedEdge, number>();
    highwayEdges.forEach(c => {
      let track = trackEndX.findIndex(endX => endX < c.xStart);
      if (track === -1) {
        track = trackEndX.length;
        trackEndX.push(c.xEnd);
      } else {
        trackEndX[track] = c.xEnd;
      }
      trackByEdge.set(c.edge, track);
    });

    const geometries: IEdgeGeometry[] = candidates.map(({ edge, a, b, needsHighway }) => {
      const track = trackByEdge.get(edge) || 0;
      const highwayY = highwayBaseY + HIGHWAY_TRACK_Y[track % HIGHWAY_TRACK_Y.length];
      const result = needsHighway ? highwayPath(a, b, highwayY) : connectorPath(a, b);
      return { edge, path: result.d, labelX: result.labelX, labelY: result.labelY };
    });
    setEdgeGeometry(geometries);
  }, [visibleEdges, orderedSteps]);

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

  const closeEditPopup = (): void => {
    setSelectedNodeId(undefined);
    setEditDraft(undefined);
  };

  const deleteSelected = (): void => {
    if (!selectedNodeId) return;
    onDeleteStep(selectedNodeId);
    closeEditPopup();
  };

  if (steps.length === 0) {
    return <div className={styles.emptyState}>No steps to show yet.</div>;
  }

  return (
    <div className={styles.canvas} ref={canvasRef}>
      <svg className={styles.edgeOverlay} ref={svgRef}>
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

      <div className={styles.grid} style={{ gridTemplateColumns: `170px repeat(${orderedSteps.length}, minmax(150px, 1fr))` }}>
        <div className={styles.corner} />
        {columnGroups.map(group => (
          <div
            className={styles.columnHeader}
            key={group.processStepId}
            style={{ gridColumn: `span ${group.stepIds.length}` }}
          >
            {drilledDownStepId
              ? (orderedSteps.find(s => s.processStepId === group.processStepId)?.processStepName || group.processStepId)
              : group.processStepId}
          </div>
        ))}

        <div className={styles.highwaySpacer} ref={highwaySpacerRef} style={{ gridColumn: '1 / -1' }} />

        {lanes.map(lane => (
          <React.Fragment key={lane}>
            <div className={styles.laneLabel}>{lane}</div>
            {orderedSteps.map(step => {
              const belongsToLane = (step.responsibleJobTitle || 'Unassigned') === lane;
              return (
                <div className={styles.laneCell} key={`${lane}-${step.id}`}>
                  {belongsToLane && (
                    <div ref={el => { if (el) nodeRefs.current.set(step.id, el); }}>
                      <ShapeNode
                        label={step.actionDescription}
                        shape={getShapeType(step)}
                        selected={selectedNodeId === step.id}
                        onClick={() => handleNodeClick(step)}
                      />
                    </div>
                  )}
                </div>
              );
            })}
          </React.Fragment>
        ))}
      </div>

      <Modal isOpen={!!(selectedNodeId && editDraft)} onDismiss={closeEditPopup} isBlocking={false} containerClassName={styles.editModal}>
        {editDraft && (
          <div className={styles.connectionsPanel}>
            <div className={styles.editHeader}>
              <h4>Edit task</h4>
              <div>
                <IconButton iconProps={{ iconName: 'Delete' }} title="Delete this task" onClick={deleteSelected} />
                <IconButton iconProps={{ iconName: 'Cancel' }} title="Close" onClick={closeEditPopup} />
              </div>
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
              <DefaultButton text="Save changes" onClick={() => { saveEdit(); closeEditPopup(); }} />
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
      </Modal>
    </div>
  );
};

export default SwimlaneCanvas;
