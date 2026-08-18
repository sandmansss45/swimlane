import * as React from 'react';
import { DefaultButton, PrimaryButton, IconButton, Modal, IDropdownOption } from '@fluentui/react';
import { toJpeg } from 'html-to-image';
import { jsPDF } from 'jspdf';
import { IProcessStep, getShapeType } from '../models/IProcessStep';
import { IRiskStatement, resolveRiskLevel, worstLinkedSeverity } from '../models/IRiskStatement';
import { IEmployee } from '../models/IEmployee';
import { IResolvedEdge, dependsOnTokensToStepIds, stepIdsToDependsOnTokens, buildDependsOnOptions } from '../utils/dependencyResolution';
import { orderStepsForTimeline, buildColumnGroups, computeDropOrder, dropKeepsDependencyOrder } from '../utils/columns';
import { connectorPath, highwayPath, pickSides, rectFromDomRect, IRect, Side } from '../utils/arrowRouting';
import ShapeNode from './shapes/ShapeNode';
import ShapeLegend from './ShapeLegend';
import ProcessStepForm, { IProcessStepFormValue } from './ProcessStepForm';
import styles from './SwimlaneCanvas.module.scss';

export interface ISwimlaneCanvasProps {
  steps: IProcessStep[]; // already filtered to the current Progress ID (and Process Step ID, if drilled down) - for display
  allSteps: IProcessStep[]; // FULL, unfiltered, original-order dataset - needed to compute row-number DependsOn tokens, which only make sense against original load order
  // This Progress ID's own steps, NOT narrowed further by drilledDownStepId
  // the way `steps` is - the "Depends on" picker's option list, since a
  // step only ever realistically depends on something in its own swimlane,
  // not one of the ~40 unrelated steps from every other flow in allSteps.
  swimlaneSteps: IProcessStep[];
  edges: IResolvedEdge[]; // resolved against the FULL, unfiltered dataset (row numbers only make sense that way) - this component only draws the ones whose endpoints are currently rendered
  riskStatements: IRiskStatement[]; // drives each shape's traffic-light fill when linked to a step
  drilledDownStepId: string | undefined;
  employees: IEmployee[];
  selectedDepartment: string | undefined;
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

// Real ResponsibleJobTitle values from the source list read as one dense
// run-on string, e.g. "Finance Manager / 6002 - Finance / United Kingdom"
// (title / employee code - department / region). Splitting it into a bold
// title line plus a quieter department+region line makes the lane header
// scannable without altering the underlying data - falls back to the raw
// string untouched for values that don't follow the pattern (e.g. a
// multi-party lane like "Vendor, Business, Finance Manager").
function formatLaneLabel(raw: string): { primary: string; secondary?: string } {
  const parts = raw.split(' / ').map(p => p.trim());
  if (parts.length !== 3) return { primary: raw };
  const [title, codeAndDept, region] = parts;
  const deptParts = codeAndDept.split(' - ');
  const dept = deptParts.length === 2 ? deptParts[1].trim() : codeAndDept;
  return { primary: title, secondary: `${dept} · ${region}` };
}

// Ported from the earlier proven vanilla-JS renderer: lanes =
// ResponsibleJobTitle, columns = Process Step ID (or per-row when drilled
// into one step), shapes per the confirmed rules, arrows measured via DOM
// refs after layout and routed via connectorPath (picks whichever side of
// each box actually faces the other node, so lines don't cut diagonally
// through unrelated boxes between them).
const SwimlaneCanvas: React.FC<ISwimlaneCanvasProps> = ({
  steps, allSteps, swimlaneSteps, edges, riskStatements, drilledDownStepId, employees, selectedDepartment, onLabelEdge, onEditStep, onDeleteStep, onMoveStep
}) => {
  const canvasRef = React.useRef<HTMLDivElement>(null);
  const svgRef = React.useRef<SVGSVGElement>(null);
  const highwaySpacerRef = React.useRef<HTMLDivElement>(null);
  const nodeRefs = React.useRef(new Map<string, HTMLDivElement>());
  const [edgeGeometry, setEdgeGeometry] = React.useState<IEdgeGeometry[]>([]);
  // Whether any edge in the CURRENT view actually needs the reserved
  // cross-lane highway strip (see the .highwaySpacer comment) - most
  // filtered views (e.g. drilled into one Progress ID with only 1-2 lanes)
  // have none, and reserving the full strip height anyway left a large
  // dead band of empty grid between the header and the first lane row.
  const [needsHighwayStrip, setNeedsHighwayStrip] = React.useState(true);
  const [selectedNodeId, setSelectedNodeId] = React.useState<string | undefined>(undefined);
  const [draftLabels, setDraftLabels] = React.useState<{ [token: string]: string }>({});
  const [editDraft, setEditDraft] = React.useState<IProcessStepFormValue | undefined>(undefined);
  const [draggingStepId, setDraggingStepId] = React.useState<string | undefined>(undefined);
  const [dragOverCellId, setDragOverCellId] = React.useState<string | undefined>(undefined);
  // Latest pointer position during a drag, kept in a ref (not state) since
  // it's read every animation frame by the auto-scroll loop below and
  // doesn't need to trigger a re-render on its own - only the scroll
  // position actually needs to move.
  const dragPointerRef = React.useRef<{ x: number; y: number } | undefined>(undefined);

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
    if (!dragged || dragged.processStepId !== columnStep.processStepId) return false;
    // Confirmed design rule: dragging can reorder within the group, but
    // never to a position that would put the step before something it
    // depends on, or after something that depends on it - that's what
    // produced backward-pointing arrows before this check existed.
    return dropKeepsDependencyOrder(allSteps, edges, draggingStepId, columnStep.id);
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

  // Native drag-and-drop doesn't auto-scroll a nested scrollable container
  // the way it scrolls the window - a real flow can run to 40 columns
  // (several screens wide when viewing "All"), so without this, a target
  // more than one screen away from the drag's starting point is simply
  // unreachable: there's no way to get the pointer there while still
  // holding the drag. Just tracks the latest pointer position here; the
  // effect below reads it every frame and does the actual scrolling, so
  // scrolling continues smoothly even while the pointer is held still
  // right at the edge, not just each time it moves.
  const handleCanvasDragOver = (e: React.DragEvent): void => {
    if (!draggingStepId) return;
    dragPointerRef.current = { x: e.clientX, y: e.clientY };
  };

  React.useEffect(() => {
    if (!draggingStepId) return;
    const EDGE = 70; // px from the canvas edge where auto-scroll kicks in
    const MAX_SPEED = 16; // px/frame at the very edge, scaling down to 0 at EDGE
    let frameId: number;

    const tick = (): void => {
      const canvas = canvasRef.current;
      const pointer = dragPointerRef.current;
      if (canvas && pointer) {
        const rect = canvas.getBoundingClientRect();
        const distLeft = pointer.x - rect.left;
        const distRight = rect.right - pointer.x;
        const distTop = pointer.y - rect.top;
        const distBottom = rect.bottom - pointer.y;
        if (distLeft < EDGE) canvas.scrollLeft -= MAX_SPEED * (1 - Math.max(0, distLeft) / EDGE);
        else if (distRight < EDGE) canvas.scrollLeft += MAX_SPEED * (1 - Math.max(0, distRight) / EDGE);
        if (distTop < EDGE) canvas.scrollTop -= MAX_SPEED * (1 - Math.max(0, distTop) / EDGE);
        else if (distBottom < EDGE) canvas.scrollTop += MAX_SPEED * (1 - Math.max(0, distBottom) / EDGE);
      }
      frameId = requestAnimationFrame(tick);
    };
    frameId = requestAnimationFrame(tick);
    return () => {
      cancelAnimationFrame(frameId);
      dragPointerRef.current = undefined;
    };
  }, [draggingStepId]);

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

  // Scoped to this swimlane's own steps, not every step across every
  // Progress ID (allSteps) - a step only ever realistically depends on
  // something in its own flow, and offering ~40 mostly-unrelated options
  // just buried the real one in noise.
  const dependsOnOptions: IDropdownOption[] = React.useMemo(
    () => buildDependsOnOptions(swimlaneSteps, selectedNodeId),
    [swimlaneSteps, selectedNodeId]
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
    //
    // The SVG is itself one of the things contributing to canvas.scrollWidth
    // (it's a real child, not removed from flow), so reading that value
    // without resetting the SVG first picks up ITS OWN previous size -
    // switching from a wide filtered view (e.g. "All") to a narrower one
    // (e.g. a single Progress ID) would otherwise never shrink back down,
    // stuck forever at the widest size ever rendered this session. Collapse
    // it first so the measurement reflects only the grid's real content.
    if (svgRef.current) {
      svgRef.current.style.width = '0px';
      svgRef.current.style.height = '0px';
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

    const highwayCandidates = candidates.filter(c => c.tier === 'highway');
    const needsHighway = highwayCandidates.length > 0;
    if (needsHighway !== needsHighwayStrip) setNeedsHighwayStrip(needsHighway);
    const highwayTracks = assignTracks(highwayCandidates);

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
  }, [visibleEdges, orderedSteps, stepsById, needsHighwayStrip]);

  React.useLayoutEffect(() => {
    // Re-runs after needsHighwayStrip flips and the spacer's own height
    // changes in the DOM, so highway edges (if any) get measured against
    // its real, settled position rather than a stale one from before the
    // resize.
    measureEdges();
    window.addEventListener('resize', measureEdges);
    return () => window.removeEventListener('resize', measureEdges);
  }, [measureEdges, needsHighwayStrip]);

  const handleNodeClick = (step: IProcessStep): void => {
    const deselecting = selectedNodeId === step.id;
    setSelectedNodeId(deselecting ? undefined : step.id);
    setDraftLabels({});
    setEditDraft(deselecting ? undefined : {
      action: step.action,
      actionDescription: step.actionDescription,
      actionType: step.actionType,
      shapeOverride: step.shapeOverride || '',
      riskLevelOverride: step.riskLevelOverride || '',
      responsibleJobTitle: step.responsibleJobTitle,
      dependsOnStepIds: dependsOnTokensToStepIds(allSteps, step.dependsOn),
      linkedRisks: step.linkedRisks || []
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
    const dependsOn = stepIdsToDependsOnTokens(allSteps, dependsOnStepIds);
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
      <div className={styles.canvas} ref={canvasRef} onDragOver={handleCanvasDragOver}>
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

      {/*
        1fr used to let each column stretch to fill whatever width the
        canvas happened to have available - fine with a full flow of 10+
        steps (they fill the space naturally), but with only 1-2 steps in a
        wide viewport it stretched each column to hundreds of pixels,
        leaving directly-connected boxes looking randomly far apart for no
        reason. Capped so a column only ever grows as far as its own
        content needs, not to fill leftover space - a short flow now just
        leaves empty canvas to the right instead of stretching.
      */}
      <div className={styles.grid} style={{ gridTemplateColumns: `170px repeat(${orderedSteps.length}, minmax(170px, 260px))` }}>
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

        <div
          className={[styles.highwaySpacer, needsHighwayStrip ? '' : styles.highwaySpacerCompact].filter(Boolean).join(' ')}
          ref={highwaySpacerRef}
          style={{ gridColumn: '1 / -1' }}
        />

        {lanes.map(lane => {
          const laneLabel = formatLaneLabel(lane);
          return (
          <React.Fragment key={lane}>
            <div className={styles.laneLabel}>
              <div className={styles.laneLabelPrimary}>{laneLabel.primary}</div>
              {laneLabel.secondary && <div className={styles.laneLabelSecondary}>{laneLabel.secondary}</div>}
            </div>
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
                        riskLevel={resolveRiskLevel(step.riskLevelOverride)}
                        linkedRiskSeverity={worstLinkedSeverity(step.linkedRisks)}
                        linkedRiskCount={(step.linkedRisks || []).length}
                        selected={selectedNodeId === step.id}
                        onClick={() => handleNodeClick(step)}
                      />
                    </div>
                  )}
                </div>
              );
            })}
          </React.Fragment>
          );
        })}
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
              <ProcessStepForm
                value={editDraft}
                onChange={setEditDraft}
                employees={employees}
                selectedDepartment={selectedDepartment}
                dependsOnOptions={dependsOnOptions}
                riskStatements={riskStatements}
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
