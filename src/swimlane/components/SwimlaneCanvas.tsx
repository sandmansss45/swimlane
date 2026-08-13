import * as React from 'react';
import { TextField, Dropdown, IDropdownOption, DefaultButton, PrimaryButton, IconButton, Modal } from '@fluentui/react';
import { toJpeg } from 'html-to-image';
import { jsPDF } from 'jspdf';
import { IProcessStep, getShapeType } from '../models/IProcessStep';
import { IRiskStatement, resolveRiskLevel } from '../models/IRiskStatement';
import { IEmployee } from '../models/IEmployee';
import { IResolvedEdge } from '../utils/dependencyResolution';
import { orderStepsForTimeline, buildColumnGroups, computeDropOrder } from '../utils/columns';
import { connectorPath, highwayPath, pickSides, rectFromDomRect, IRect, Side } from '../utils/arrowRouting';
import ShapeNode from './shapes/ShapeNode';
import ShapeLegend from './ShapeLegend';
import EmployeePicker from './EmployeePicker';
import styles from './SwimlaneCanvas.module.scss';

const SHAPE_OPTIONS: IDropdownOption[] = [
  { key: '', text: '(use Action Type / wording heuristic)' },
  { key: 'Process Step', text: 'Process (rounded rectangle)' },
  { key: 'Decision', text: 'Decision (diamond)' },
  { key: 'Approval', text: 'Approval (circle)' },
  { key: 'Document', text: 'Document (artifact)' }
];

const RISK_LEVEL_OPTIONS: IDropdownOption[] = [
  { key: '', text: '(use Risk Register)' },
  { key: 'High', text: 'High' },
  { key: 'Medium', text: 'Medium' },
  { key: 'Low', text: 'Low' }
];

const truncate = (text: string, max: number): string =>
  text.length > max ? `${text.slice(0, max - 1)}…` : text;

export interface ISwimlaneCanvasProps {
  steps: IProcessStep[]; // already filtered to the current Progress ID (and Process Step ID, if drilled down) - for display
  allSteps: IProcessStep[]; // FULL, unfiltered, original-order dataset - needed to compute row-number DependsOn tokens and to offer every step as a "depends on" option regardless of what's currently filtered into view
  edges: IResolvedEdge[]; // resolved against the FULL, unfiltered dataset (row numbers only make sense that way) - this component only draws the ones whose endpoints are currently rendered
  riskStatements: IRiskStatement[]; // drives each shape's traffic-light fill when linked to a step
  drilledDownStepId: string | undefined;
  employees: IEmployee[];
  selectedRegion: string | undefined;
  onLabelEdge: (toRowId: string, token: string, label: string) => void;
  onEditStep: (step: IProcessStep) => void;
  onDeleteStep: (stepId: string) => void;
  // Drag-and-drop: dropping a step onto another step's cell moves it to
  // that lane and re-sequences it to sit right after that step within
  // their shared Process Step ID group - see computeDropOrder for why
  // dragging is confined to one group.
  onMoveStep: (updated: IProcessStep) => void;
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
  riskLevelOverride: string;
  responsibleJobTitle: string;
  dependsOnStepIds: string[];
}

const SwimlaneCanvas: React.FC<ISwimlaneCanvasProps> = ({
  steps, allSteps, edges, riskStatements, drilledDownStepId, employees, selectedRegion, onLabelEdge, onEditStep, onDeleteStep, onMoveStep
}) => {
  const canvasRef = React.useRef<HTMLDivElement>(null);
  const svgRef = React.useRef<SVGSVGElement>(null);
  const highwaySpacerRef = React.useRef<HTMLDivElement>(null);
  const nodeRefs = React.useRef(new Map<string, HTMLDivElement>());
  const [edgeGeometry, setEdgeGeometry] = React.useState<IEdgeGeometry[]>([]);
  const [selectedNodeId, setSelectedNodeId] = React.useState<string | undefined>(undefined);
  const [draftLabels, setDraftLabels] = React.useState<{ [token: string]: string }>({});
  const [editDraft, setEditDraft] = React.useState<IEditDraft | undefined>(undefined);
  const [draggingStepId, setDraggingStepId] = React.useState<string | undefined>(undefined);
  const [dragOverCellId, setDragOverCellId] = React.useState<string | undefined>(undefined);

  const stepsById = React.useMemo(() => new Map(steps.map(s => [s.id, s])), [steps]);

  // Drag-and-drop: `columnStep` is whichever step the target COLUMN
  // belongs to (orderedSteps[i] for that column, regardless of which
  // lane's cell is actually being dragged over) - a cell can be empty
  // (its lane doesn't own that column's step) and still be a valid,
  // meaningful drop target, since dropping there both re-sequences the
  // dragged step relative to columnStep AND reassigns it to this cell's
  // lane.
  const isValidDropTarget = (columnStep: IProcessStep): boolean => {
    if (!draggingStepId || draggingStepId === columnStep.id) return false;
    const dragged = stepsById.get(draggingStepId);
    return !!dragged && dragged.processStepId === columnStep.processStepId;
  };

  const handleDragStart = (step: IProcessStep) => (e: React.DragEvent): void => {
    setDraggingStepId(step.id);
    e.dataTransfer.effectAllowed = 'move';
    // Firefox refuses to start a drag at all unless setData is called.
    e.dataTransfer.setData('text/plain', step.id);
  };

  const handleDragEnd = (): void => {
    setDraggingStepId(undefined);
    setDragOverCellId(undefined);
  };

  const handleCellDragOver = (columnStep: IProcessStep, lane: string) => (e: React.DragEvent): void => {
    if (!isValidDropTarget(columnStep)) return;
    e.preventDefault(); // only opt into "droppable" when valid - otherwise leave the browser's own "not allowed" cursor
    e.dataTransfer.dropEffect = 'move';
    setDragOverCellId(`${lane}-${columnStep.id}`);
  };

  const handleCellDragLeave = (lane: string, columnStep: IProcessStep) => (): void => {
    setDragOverCellId(prev => (prev === `${lane}-${columnStep.id}` ? undefined : prev));
  };

  const handleDrop = (columnStep: IProcessStep, lane: string) => (e: React.DragEvent): void => {
    e.preventDefault();
    setDragOverCellId(undefined);
    if (!draggingStepId || !isValidDropTarget(columnStep)) { setDraggingStepId(undefined); return; }
    const dragged = stepsById.get(draggingStepId);
    if (!dragged) { setDraggingStepId(undefined); return; }
    const newOrder = computeDropOrder(allSteps, draggingStepId, columnStep.id);
    onMoveStep({
      ...dragged,
      responsibleJobTitle: lane === 'Unassigned' ? '' : lane,
      manualOrder: newOrder !== undefined ? newOrder : dragged.manualOrder
    });
    setDraggingStepId(undefined);
  };

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
    // Parallel tracks inside the reserved cross-lane strip. Two
    // independent edges that both genuinely need it can still land on
    // overlapping X-ranges - give them real separation so they don't read
    // as one tangled, doubled-up line.
    const HIGHWAY_TRACK_Y = [0, -24, 24, -48, 48];
    // Parallel tracks for a same-lane "local hop" - these only ever share
    // space with OTHER hops in that same row, which is rare, so a
    // tighter spread is enough and keeps the bend close to the row.
    const LOCAL_HOP_TRACK_Y = [0, -12, 12];
    // How far above a lane's tallest box top a hop clears before bending
    // sideways - see the row-gap comment in SwimlaneCanvas.module.scss for
    // why 44px of gap comfortably fits this even against the tightest
    // shape (the 140px approval circle).
    const LOCAL_HOP_CLEARANCE = 22;

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

    // Topmost edge of each lane's tallest rendered box (a diamond sits
    // much higher within its own cell than a rect does) - a local hop
    // clears everything in its row by bending above THIS, not just above
    // its own two endpoints, so it can't clip a shorter sibling that
    // happens to sit between them.
    const laneMinTop = new Map<string, number>();
    orderedSteps.forEach(step => {
      const rect = rectById.get(step.id);
      if (!rect) return;
      const lane = step.responsibleJobTitle || 'Unassigned';
      const current = laneMinTop.get(lane);
      if (current === undefined || rect.top < current) laneMinTop.set(lane, rect.top);
    });

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

    // Three routing tiers, cheapest first:
    //  - direct: nothing between the two boxes, draw straight/dogleg.
    //  - localHop: blocked, but only by a SIBLING IN THE SAME LANE (same
    //    row, some column between them) - decision branches reconverging
    //    around a sibling branch is exactly this, and it's the overwhelming
    //    majority of "blocked" edges in a real swimlane. A hop confined to
    //    that one row clears it without ever leaving the row.
    //  - highway: blocked by something in a DIFFERENT lane - the straight
    //    path crosses other rows too, which a hop confined to one row can't
    //    route around, so it takes the shared cross-lane strip at the top.
    type RoutingTier = 'direct' | 'localHop' | 'highway';
    interface ICandidate { edge: IResolvedEdge; a: IRect; b: IRect; tier: RoutingTier; lane?: string }
    const candidates: ICandidate[] = [];

    visibleEdges.forEach(edge => {
      const a = rectById.get(edge.fromRowId);
      const b = rectById.get(edge.toRowId);
      if (!a || !b) return;
      if (pathIsClear(edge.fromRowId, a, edge.toRowId, b)) {
        candidates.push({ edge, a, b, tier: 'direct' });
        return;
      }
      const fromLane = stepsById.get(edge.fromRowId)?.responsibleJobTitle || 'Unassigned';
      const toLane = stepsById.get(edge.toRowId)?.responsibleJobTitle || 'Unassigned';
      if (fromLane === toLane) {
        candidates.push({ edge, a, b, tier: 'localHop', lane: fromLane });
      } else {
        candidates.push({ edge, a, b, tier: 'highway' });
      }
    });

    // Classic greedy interval-coloring (sort by start, reuse the first
    // track that's already clear by then) so two connections whose
    // horizontal spans overlap land on distinct parallel Y-levels instead
    // of drawing on top of each other. Run separately per group - the
    // shared cross-lane strip is one group, and each lane gets its own
    // independent local-hop group, since hops in different lanes never
    // share visual space and so never need to be coordinated together.
    const assignTracks = (group: ICandidate[]): Map<IResolvedEdge, number> => {
      const sorted = group
        .map(c => ({ c, xStart: Math.min(c.a.cx, c.b.cx), xEnd: Math.max(c.a.cx, c.b.cx) }))
        .sort((x, y) => x.xStart - y.xStart);
      const trackEndX: number[] = [];
      const trackByEdge = new Map<IResolvedEdge, number>();
      sorted.forEach(({ c, xStart, xEnd }) => {
        let track = trackEndX.findIndex(endX => endX < xStart);
        if (track === -1) {
          track = trackEndX.length;
          trackEndX.push(xEnd);
        } else {
          trackEndX[track] = xEnd;
        }
        trackByEdge.set(c.edge, track);
      });
      return trackByEdge;
    };

    const highwayTracks = assignTracks(candidates.filter(c => c.tier === 'highway'));

    const localHopGroups = new Map<string, ICandidate[]>();
    candidates.filter(c => c.tier === 'localHop').forEach(c => {
      const key = c.lane || 'Unassigned';
      const list = localHopGroups.get(key) || [];
      list.push(c);
      localHopGroups.set(key, list);
    });
    const localHopTracks = new Map<IResolvedEdge, number>();
    localHopGroups.forEach(group => {
      assignTracks(group).forEach((track, edge) => localHopTracks.set(edge, track));
    });

    // Multiple edges attaching to the same side of the same box would
    // otherwise all pass through that side's exact center point - two
    // arrows entering/exiting through the same hole, which is what made a
    // decision's branch and an unrelated reconvergence edge look fused
    // together right where they neared the same box. Group every edge's
    // endpoint by (node, side) regardless of whether it's the source or
    // target end there, sort by where the OTHER end of each edge sits (so
    // the spread reads left-to-right sensibly instead of crossing more
    // than it has to), then hand out symmetric offsets around the center.
    interface IAttachment { edge: IResolvedEdge; role: 'from' | 'to'; side: Side; otherCx: number; otherCy: number }
    const attachments: IAttachment[] = [];
    candidates.forEach(c => {
      const sides = c.tier === 'direct'
        ? pickSides(c.a, c.b)
        : { fromSide: 'top' as Side, toSide: 'top' as Side }; // localHop/highway always attach from the top
      attachments.push({ edge: c.edge, role: 'from', side: sides.fromSide, otherCx: c.b.cx, otherCy: c.b.cy });
      attachments.push({ edge: c.edge, role: 'to', side: sides.toSide, otherCx: c.a.cx, otherCy: c.a.cy });
    });

    const attachGroups = new Map<string, IAttachment[]>();
    attachments.forEach(att => {
      const nodeId = att.role === 'from' ? att.edge.fromRowId : att.edge.toRowId;
      const key = `${nodeId}::${att.side}`;
      const list = attachGroups.get(key) || [];
      list.push(att);
      attachGroups.set(key, list);
    });

    const fromOffset = new Map<IResolvedEdge, number>();
    const toOffset = new Map<IResolvedEdge, number>();
    attachGroups.forEach(group => {
      if (group.length < 2) return;
      const onXAxis = group[0].side === 'top' || group[0].side === 'bottom';
      const sorted = [...group].sort((p, q) => (onXAxis ? p.otherCx - q.otherCx : p.otherCy - q.otherCy));
      const step = onXAxis ? 26 : 16; // top/bottom spreads along a ~150-190px-wide box; left/right along a shorter side
      const maxAbs = onXAxis ? 65 : 40;
      const start = -((sorted.length - 1) * step) / 2;
      sorted.forEach((att, i) => {
        const offset = Math.max(-maxAbs, Math.min(maxAbs, start + i * step));
        (att.role === 'from' ? fromOffset : toOffset).set(att.edge, offset);
      });
    });

    const geometries: IEdgeGeometry[] = candidates.map(({ edge, a, b, tier, lane }) => {
      const aOffset = fromOffset.get(edge) || 0;
      const bOffset = toOffset.get(edge) || 0;
      if (tier === 'direct') {
        const result = connectorPath(a, b, aOffset, bOffset);
        return { edge, path: result.d, labelX: result.labelX, labelY: result.labelY };
      }
      if (tier === 'localHop') {
        const track = localHopTracks.get(edge) || 0;
        const laneTop = laneMinTop.get(lane || 'Unassigned');
        const baseTop = laneTop === undefined ? Math.min(a.top, b.top) : laneTop;
        const hopY = baseTop - LOCAL_HOP_CLEARANCE + LOCAL_HOP_TRACK_Y[track % LOCAL_HOP_TRACK_Y.length];
        const result = highwayPath(a, b, hopY, aOffset, bOffset);
        return { edge, path: result.d, labelX: result.labelX, labelY: result.labelY };
      }
      const track = highwayTracks.get(edge) || 0;
      const highwayY = highwayBaseY + HIGHWAY_TRACK_Y[track % HIGHWAY_TRACK_Y.length];
      const result = highwayPath(a, b, highwayY, aOffset, bOffset);
      return { edge, path: result.d, labelX: result.labelX, labelY: result.labelY };
    });
    setEdgeGeometry(geometries);
  }, [visibleEdges, orderedSteps, stepsById]);

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
      riskLevelOverride: step.riskLevelOverride || '',
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

  const [exporting, setExporting] = React.useState(false);

  const exportToPdf = async (): Promise<void> => {
    const canvas = canvasRef.current;
    if (!canvas || exporting) return;
    setExporting(true);
    try {
      // The live element only shows its scrolled/visible portion - render
      // the capture at the FULL scrollable content size instead (same
      // scrollWidth/scrollHeight the SVG overlay is already sized to), and
      // override overflow so nothing gets clipped in the snapshot the way
      // it would in the live, scrolled view.
      const width = canvas.scrollWidth;
      const height = canvas.scrollHeight;
      // PNG at pixelRatio 2 on a wide diagram produced a ~95MB file
      // (lossless full-resolution bitmap) - completely impractical to
      // download or email. JPEG at high quality and native resolution
      // brings a diagram this size down to single-digit MB while staying
      // sharp enough to read every label.
      const dataUrl = await toJpeg(canvas, {
        width,
        height,
        pixelRatio: 1,
        quality: 0.92,
        backgroundColor: '#ffffff',
        style: { overflow: 'visible', maxHeight: 'none' }
      });
      const pdf = new jsPDF({
        orientation: width >= height ? 'landscape' : 'portrait',
        unit: 'px',
        format: [width, height]
      });
      pdf.addImage(dataUrl, 'JPEG', 0, 0, width, height);
      pdf.save(`swimlane-studio-${new Date().toISOString().slice(0, 10)}.pdf`);
    } catch (err) {
      // eslint-disable-next-line no-console
      console.error('PDF export failed', err);
    } finally {
      setExporting(false);
    }
  };

  if (steps.length === 0) {
    return <div className={styles.emptyState}>No steps to show yet.</div>;
  }

  return (
    <>
      <div className={styles.canvasToolbar}>
        <DefaultButton
          text={exporting ? 'Exporting…' : 'Export to PDF'}
          iconProps={{ iconName: 'PDF' }}
          onClick={exportToPdf}
          disabled={exporting}
        />
      </div>
      <div className={styles.canvas} ref={canvasRef}>
      <ShapeLegend />
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
              const cellKey = `${lane}-${step.id}`;
              const cellClassName = [
                styles.laneCell,
                isValidDropTarget(step) ? styles.validDropTarget : '',
                dragOverCellId === cellKey ? styles.dragOver : ''
              ].filter(Boolean).join(' ');
              return (
                <div
                  className={cellClassName}
                  key={cellKey}
                  onDragOver={handleCellDragOver(step, lane)}
                  onDragLeave={handleCellDragLeave(lane, step)}
                  onDrop={handleDrop(step, lane)}
                >
                  {belongsToLane && (
                    <div
                      ref={el => { if (el) nodeRefs.current.set(step.id, el); }}
                      className={[styles.draggableNode, draggingStepId === step.id ? styles.dragging : ''].filter(Boolean).join(' ')}
                      draggable
                      onDragStart={handleDragStart(step)}
                      onDragEnd={handleDragEnd}
                    >
                      <ShapeNode
                        label={step.actionDescription}
                        shape={getShapeType(step)}
                        riskLevel={resolveRiskLevel(step, riskStatements)}
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
              <Dropdown
                label="Risk level"
                selectedKey={editDraft.riskLevelOverride}
                options={RISK_LEVEL_OPTIONS}
                onChange={(_e, option) => setEditDraft(prev => (prev && option ? { ...prev, riskLevelOverride: String(option.key) } : prev))}
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
              <PrimaryButton text="Save changes" onClick={() => { saveEdit(); closeEditPopup(); }} />
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
    </>
  );
};

export default SwimlaneCanvas;
