import * as React from 'react';
import { DefaultButton, PrimaryButton, IconButton, Modal, IDropdownOption, Callout, TextField } from '@fluentui/react';
import { toJpeg } from 'html-to-image';
import { jsPDF } from 'jspdf';
import { IProcessStep, getShapeType } from '../models/IProcessStep';
import { IRiskStatement, worstLinkedSeverity } from '../models/IRiskStatement';
import { IEmployee } from '../models/IEmployee';
import { SwimlaneStage } from '../models/ISwimlaneStatus';
import { IResolvedEdge, dependsOnTokensToStepIds, stepIdsToDependsOnTokens, buildDependsOnOptions } from '../utils/dependencyResolution';
import { orderStepsForTimeline, buildColumnGroups, computeDropOrder, dropKeepsDependencyOrder } from '../utils/columns';
import { connectorPath, highwayPath, pickSides, rectFromDomRect, IRect, Side } from '../utils/arrowRouting';
import ShapeNode from './shapes/ShapeNode';
import ShapeLegend from './ShapeLegend';
import ProcessStepForm, { IProcessStepFormValue } from './ProcessStepForm';
import styles from './SwimlaneCanvas.module.scss';

export interface ISwimlaneCanvasProps {
  steps: IProcessStep[]; // already filtered to the current Process ID (and Process Step ID, if drilled down) - for display
  allSteps: IProcessStep[]; // FULL, unfiltered, original-order dataset - needed to compute row-number DependsOn tokens, which only make sense against original load order
  // This Process ID's own steps, NOT narrowed further by drilledDownStepId
  // the way `steps` is - the "Depends on" picker's option list, since a
  // step only ever realistically depends on something in its own swimlane,
  // not one of the ~40 unrelated steps from every other flow in allSteps.
  swimlaneSteps: IProcessStep[];
  edges: IResolvedEdge[]; // resolved against the FULL, unfiltered dataset (row numbers only make sense that way) - this component only draws the ones whose endpoints are currently rendered
  riskStatements: IRiskStatement[]; // drives each shape's traffic-light fill when linked to a step
  drilledDownStepId: string | undefined;
  employees: IEmployee[];
  // True while the swimlane currently on screen is locked (see
  // IProcessIdLock) - disables dragging entirely and switches the edit
  // panel to read-only (still opens, just can't Save or Delete), rather
  // than hiding the panel altogether - people should still be able to
  // look at a locked step's details.
  isLocked: boolean;
  onLabelEdge: (toRowId: string, token: string, label: string) => void;
  onEditStep: (step: IProcessStep) => void;
  onDeleteStep: (stepId: string) => void;
  // Drag-and-drop: dropping a step onto another step's cell moves it to
  // that lane and re-sequences it to sit right after that step within
  // their shared Process Step ID group - see computeDropOrder for why
  // dragging is confined to one group.
  onMoveStep: (updated: IProcessStep) => void;
  // Drag-from-legend: dropping a shape from ShapeLegend onto a cell
  // creates a brand new step there - lane and Process Step ID come from
  // whatever cell it landed on (see computeInsertOrderAfter), shape from
  // which legend swatch was dragged. Everything else starts blank/default,
  // filled in via the edit panel that opens automatically right after
  // (see autoOpenStepId).
  onCreateStep: (shapeOverride: string, lane: string, columnStep: IProcessStep) => void;
  // One-shot signal from the parent: once set, opens this step's edit
  // panel exactly as if it had been clicked, then the parent should clear
  // it back to undefined via onAutoOpenHandled. Used right after
  // onCreateStep so a freshly dropped step's details can be filled in
  // immediately, without a separate click to find and open it.
  autoOpenStepId?: string;
  onAutoOpenHandled?: () => void;
  // Draft/Finalised status for the swimlane currently on screen (see
  // ISwimlaneStatus) - purely a display label with a toggle button,
  // shown underneath ShapeLegend. No relation to isLocked/onEditStep -
  // this never blocks anything, it's informational only.
  swimlaneStage: SwimlaneStage;
  // Undefined means no one has ever explicitly set a stage - the current
  // Draft default has no attribution to show.
  stageSetBy: string | undefined;
  onToggleStage: () => void;
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
  steps, allSteps, swimlaneSteps, edges, riskStatements, drilledDownStepId, employees, isLocked, onLabelEdge, onEditStep, onDeleteStep, onMoveStep,
  onCreateStep, autoOpenStepId, onAutoOpenHandled, swimlaneStage, stageSetBy, onToggleStage
}) => {
  const canvasRef = React.useRef<HTMLDivElement>(null);
  const svgRef = React.useRef<SVGSVGElement>(null);
  const highwaySpacerRef = React.useRef<HTMLDivElement>(null);
  const nodeRefs = React.useRef(new Map<string, HTMLDivElement>());
  const [edgeGeometry, setEdgeGeometry] = React.useState<IEdgeGeometry[]>([]);
  // Whether any edge in the CURRENT view actually needs the reserved
  // cross-lane highway strip (see the .highwaySpacer comment) - most
  // filtered views (e.g. drilled into one Process ID with only 1-2 lanes)
  // have none, and reserving the full strip height anyway left a large
  // dead band of empty grid between the header and the first lane row.
  const [needsHighwayStrip, setNeedsHighwayStrip] = React.useState(true);
  const [selectedNodeId, setSelectedNodeId] = React.useState<string | undefined>(undefined);
  const [draftLabels, setDraftLabels] = React.useState<{ [token: string]: string }>({});
  // Set while the small "Label or delete this dependency" Callout is open
  // for one specific arrow (see the edge hit-area's onClick) - point is
  // the click's viewport coordinates, since there's no single stable DOM
  // element per edge to anchor a Callout's target to the way there would
  // be for an ordinary button.
  const [edgePopup, setEdgePopup] = React.useState<{ edge: IResolvedEdge; point: { x: number; y: number } } | undefined>(undefined);
  const [editDraft, setEditDraft] = React.useState<IProcessStepFormValue | undefined>(undefined);
  const [draggingStepId, setDraggingStepId] = React.useState<string | undefined>(undefined);
  // Set while dragging a shape IN from ShapeLegend rather than moving an
  // existing step already on the canvas - the two are mutually exclusive
  // (only one drag can be in progress at once) but kept as separate state
  // rather than one union, since they have almost entirely different
  // validity rules and drop handling (see isValidDropTargetForCurrentDrag/
  // handleDrop below).
  const [draggingNewShape, setDraggingNewShape] = React.useState<string | undefined>(undefined);
  // Set while dragging a step's connector handle (see ShapeNode) to draw a
  // new dependency link onto another step - a third, independent drag
  // mode alongside draggingStepId (move) and draggingNewShape (create),
  // never active at the same time as either since only one native drag
  // gesture can be in progress at once.
  const [connectingFromStepId, setConnectingFromStepId] = React.useState<string | undefined>(undefined);
  const [dragOverCellId, setDragOverCellId] = React.useState<string | undefined>(undefined);
  // Which half of dragOverCellId's cell the pointer is currently over -
  // left half means the dragged step would land immediately BEFORE
  // columnStep, right half immediately AFTER (see computeDropOrder).
  // Only meaningful together with dragOverCellId; stale otherwise.
  const [dragOverPosition, setDragOverPosition] = React.useState<'before' | 'after'>('after');
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
  // Coarse check: is this column even a candidate drop target at all,
  // regardless of which half of the cell (before/after) ends up hovered -
  // used for the dashed "possible target" outline shown on every
  // candidate cell for the whole drag, before the pointer has actually
  // reached any specific one. The precise, side-aware check that decides
  // whether a specific half is ACTUALLY droppable right now lives in
  // handleCellDragOver/handleDrop below.
  const isValidDropTarget = (columnStep: IProcessStep): boolean => {
    if (isLocked || !draggingStepId || draggingStepId === columnStep.id) return false;
    const dragged = stepsById.get(draggingStepId);
    if (!dragged || dragged.processStepId !== columnStep.processStepId) return false;
    // Confirmed design rule: dragging can reorder within the group, but
    // never to a position that would put the step before something it
    // depends on, or after something that depends on it - that's what
    // produced backward-pointing arrows before this check existed.
    // Either side counts here - a cell showing the dashed outline just
    // means SOME drop there is legal, not that both halves necessarily are.
    return dropKeepsDependencyOrder(allSteps, edges, draggingStepId, columnStep.id, true)
      || dropKeepsDependencyOrder(allSteps, edges, draggingStepId, columnStep.id, false);
  };

  // A brand new step (dragged in from ShapeLegend) has no dependency
  // relationships yet and no "self" to exclude, so unlike
  // isValidDropTarget above, EVERY column is a valid target once the
  // swimlane isn't locked - it'll simply adopt whatever columnStep's
  // Process Step ID and lane it lands on.
  const isValidNewShapeDropTarget = (): boolean => !isLocked && !!draggingNewShape;

  const isValidDropTargetForCurrentDrag = (columnStep: IProcessStep): boolean =>
    draggingNewShape ? isValidNewShapeDropTarget() : isValidDropTarget(columnStep);

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
    if (draggingNewShape) {
      if (!isValidNewShapeDropTarget()) return;
      e.preventDefault(); // only opt into "droppable" when valid - otherwise leave the browser's own "not allowed" cursor
      e.dataTransfer.dropEffect = 'copy';
      setDragOverCellId(`${lane}-${columnStep.id}`);
      return;
    }
    if (!draggingStepId || draggingStepId === columnStep.id) return;
    const dragged = stepsById.get(draggingStepId);
    if (!dragged || dragged.processStepId !== columnStep.processStepId || isLocked) return;
    // Left half of the cell = drop before columnStep, right half = after -
    // the only way to land a step at the very front of its group, which
    // dropping-always-after (the original, only behaviour) could never
    // do, since there's no column further left than the first one to
    // "drop after" to get the same result.
    const rect = e.currentTarget.getBoundingClientRect();
    const insertBefore = e.clientX - rect.left < rect.width / 2;
    if (!dropKeepsDependencyOrder(allSteps, edges, draggingStepId, columnStep.id, insertBefore)) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    setDragOverCellId(`${lane}-${columnStep.id}`);
    setDragOverPosition(insertBefore ? 'before' : 'after');
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
    if (!draggingStepId && !draggingNewShape && !connectingFromStepId) return;
    dragPointerRef.current = { x: e.clientX, y: e.clientY };
  };

  React.useEffect(() => {
    if (!draggingStepId && !draggingNewShape && !connectingFromStepId) return;
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
  }, [draggingStepId, draggingNewShape, connectingFromStepId]);

  const handleDrop = (columnStep: IProcessStep, lane: string) => (e: React.DragEvent): void => {
    e.preventDefault();
    setDragOverCellId(undefined);
    if (draggingNewShape) {
      const shapeOverride = draggingNewShape;
      setDraggingNewShape(undefined);
      if (!isValidNewShapeDropTarget()) return;
      onCreateStep(shapeOverride, lane, columnStep);
      return;
    }
    if (!draggingStepId) { setDraggingStepId(undefined); return; }
    const dragged = stepsById.get(draggingStepId);
    if (!dragged) { setDraggingStepId(undefined); return; }
    // Recomputed from the drop event's own coordinates rather than
    // trusting the last dragOverPosition state - the drop can land on a
    // different element than the last dragover fired on (e.g. a fast
    // pointer move), so this is the one guaranteed-fresh read of exactly
    // where it actually landed.
    const rect = e.currentTarget.getBoundingClientRect();
    const insertBefore = e.clientX - rect.left < rect.width / 2;
    if (dragged.processStepId !== columnStep.processStepId
      || !dropKeepsDependencyOrder(allSteps, edges, draggingStepId, columnStep.id, insertBefore)) {
      setDraggingStepId(undefined);
      return;
    }
    const newOrder = computeDropOrder(allSteps, draggingStepId, columnStep.id, insertBefore);
    onMoveStep({
      ...dragged,
      responsibleJobTitle: lane === 'Unassigned' ? '' : lane,
      manualOrder: newOrder !== undefined ? newOrder : dragged.manualOrder
    });
    setDraggingStepId(undefined);
  };

  // Scoped to this swimlane's own steps, not every step across every
  // Process ID (allSteps) - a step only ever realistically depends on
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

  // Drag-to-connect: unlike isValidDropTarget above, a link can legally
  // cross Process Step ID groups within the same swimlane (DependsOn
  // already allows that - see buildDependsOnOptions), so group membership
  // isn't the constraint here. What IS still enforced, matching the same
  // "arrows only move forward" rule the move-drag already follows: the
  // target has to sit AT OR AFTER the source in the current left-to-right
  // column order, or the new arrow would point backward. Comparing
  // positions in orderedSteps (the same array that defines the columns
  // themselves) makes this correct across group boundaries too, and
  // rules out circular links as a side effect - a genuine cycle would
  // need at least one backward edge, which this already rejects.
  const isValidConnectionTarget = (targetStep: IProcessStep): boolean => {
    if (isLocked || !connectingFromStepId || connectingFromStepId === targetStep.id) return false;
    const sourceIdx = orderedSteps.findIndex(s => s.id === connectingFromStepId);
    const targetIdx = orderedSteps.findIndex(s => s.id === targetStep.id);
    if (sourceIdx === -1 || targetIdx === -1) return false;
    return targetIdx > sourceIdx;
  };

  const handleConnectDragStart = (step: IProcessStep) => (): void => {
    setConnectingFromStepId(step.id);
  };

  const handleConnectDragEnd = (): void => {
    setConnectingFromStepId(undefined);
  };

  // Only intercepts the event while a connection drag is actually in
  // progress - otherwise returns without calling preventDefault/
  // stopPropagation, so the event bubbles up untouched to the parent
  // cell's own onDragOver (the move/create-from-legend handlers), which
  // still need to see it when THIS drag mode isn't the one active.
  const handleConnectDragOver = (targetStep: IProcessStep) => (e: React.DragEvent): void => {
    if (!connectingFromStepId || !isValidConnectionTarget(targetStep)) return;
    e.preventDefault();
    e.stopPropagation();
  };

  const handleConnectDrop = (targetStep: IProcessStep) => (e: React.DragEvent): void => {
    if (!connectingFromStepId) return; // not our drag - let the cell's own onDrop handle it
    e.preventDefault();
    e.stopPropagation();
    const sourceId = connectingFromStepId;
    setConnectingFromStepId(undefined);
    if (!isValidConnectionTarget(targetStep)) return;
    // Adds to whatever the target already depends on, rather than
    // replacing it - a step can have more than one predecessor, same as
    // setting this by hand in the edit panel's Depends on field.
    // Silently no-ops if the link already exists instead of duplicating
    // the token.
    const existingStepIds = dependsOnTokensToStepIds(allSteps, targetStep.dependsOn);
    if (existingStepIds.includes(sourceId)) return;
    onEditStep({ ...targetStep, dependsOn: stepIdsToDependsOnTokens(allSteps, [...existingStepIds, sourceId]) });
  };

  // Only edges whose both ends are currently rendered - the rest belong to
  // a different Process ID / Process Step ID that isn't in view right now.
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
    // (e.g. a single Process ID) would otherwise never shrink back down,
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

    // A direct cross-lane edge that happens to have a clear straight-line
    // path can still land on a box that OTHER edges already reach via the
    // highway/local-hop system - confirmed real user report: a
    // diagonally-routed "yes" edge cutting straight across visibly
    // crossed/tangled with unrelated highway traffic converging on the
    // same box from a completely different side. pathIsClear only checks
    // for STEP BOXES in the way, never other ALREADY-ROUTED EDGES, so
    // this diagonal had no way to know it was about to cut through a busy
    // convergence point. Upgrade it to highway too, so it joins the same
    // coordinated track system (assignTracks below) instead of drawing
    // independently through space other edges already occupy. Same-lane
    // direct edges are excluded - they're short/local and never reach
    // anywhere near the highway convergence zone in the first place.
    const nodesWithConvergingTraffic = new Set<string>();
    candidates.forEach(c => {
      if (c.tier === 'highway' || c.tier === 'localHop') {
        nodesWithConvergingTraffic.add(c.edge.fromRowId);
        nodesWithConvergingTraffic.add(c.edge.toRowId);
      }
    });
    candidates.forEach(c => {
      if (c.tier !== 'direct') return;
      const fromLane = stepsById.get(c.edge.fromRowId)?.responsibleJobTitle || 'Unassigned';
      const toLane = stepsById.get(c.edge.toRowId)?.responsibleJobTitle || 'Unassigned';
      if (fromLane === toLane) return;
      if (nodesWithConvergingTraffic.has(c.edge.fromRowId) || nodesWithConvergingTraffic.has(c.edge.toRowId)) {
        c.tier = 'highway';
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
      responsibleJobTitle: step.responsibleJobTitle,
      dependsOnStepIds: dependsOnTokensToStepIds(allSteps, step.dependsOn),
      linkedRisks: step.linkedRisks || [],
      sopLink: step.sopLink || '',
      delegationOfAuthorityLink: step.delegationOfAuthorityLink || ''
    });
  };

  // Opens a just-created step's edit panel automatically, right after
  // dropping a shape from the legend (see onCreateStep) - the same panel
  // a click would open, just triggered by the parent instead of a second,
  // separate click to go find the new step and open it by hand. Waits for
  // `steps` to actually contain it (the parent creates it asynchronously
  // via addProcessStep) rather than opening on the stale, pre-creation
  // props from the same render the drop happened in.
  React.useEffect(() => {
    if (!autoOpenStepId) return;
    const created = steps.find(s => s.id === autoOpenStepId);
    if (!created) return;
    handleNodeClick(created);
    onAutoOpenHandled?.();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [autoOpenStepId, steps]);

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

  // Removes just this one dependency - offered from the small popup
  // clicking an arrow opens (see edgePopup), no need to open the source
  // step's edit panel and find it in the Outgoing connections list first.
  // Goes through onEditStep, the same path every other edit takes (not a
  // new action type), so it gets undo for free (see handleEditStep/
  // lastAction in SwimlaneStudio.tsx) - consistent with how this app
  // already treats "editing" a link elsewhere (RiskLinkPicker: remove and
  // re-add, never in-place edit).
  const deleteEdge = (edge: IResolvedEdge): void => {
    if (isLocked) return;
    const target = stepsById.get(edge.toRowId);
    if (!target) return;
    onEditStep({ ...target, dependsOn: target.dependsOn.filter(token => token !== edge.token) });
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
    // setExporting(true) only SCHEDULES the re-render that hides empty-cell
    // chrome and the connector handles (see cellClassName/onConnectorDragStart
    // below) - without waiting for it to actually paint, toJpeg would
    // capture the DOM as it looked a frame ago, before either change took
    // effect. Two rAFs (not one) reliably lands after a real paint, not
    // just after the next scheduled frame callback.
    await new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
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
      <ShapeLegend
        onDragShapeStart={shapeOverride => !isLocked && setDraggingNewShape(shapeOverride)}
        onDragShapeEnd={() => setDraggingNewShape(undefined)}
        stage={swimlaneStage}
        stageSetBy={stageSetBy}
        onToggleStage={onToggleStage}
      />
      <svg className={styles.edgeOverlay} ref={svgRef}>
        <defs>
          <marker id="swimlaneArrowhead" markerWidth="9" markerHeight="9" refX="7" refY="4.5" orient="auto">
            <path d="M0,0 L9,4.5 L0,9 Z" fill="#3c4a63" />
          </marker>
          <marker id="swimlaneArrowheadHighlighted" markerWidth="9" markerHeight="9" refX="7" refY="4.5" orient="auto">
            <path d="M0,0 L9,4.5 L0,9 Z" fill="#1441b9" />
          </marker>
        </defs>
        {edgeGeometry.map(({ edge, path, labelX, labelY }) => {
          const highlighted = selectedNodeId !== undefined && edge.fromRowId === selectedNodeId;
          return (
            <g key={`${edge.fromRowId}-${edge.toRowId}-${edge.token}`}>
              {/*
                Wide, invisible hit area, rendered BEFORE (so it paints
                underneath) the visible line below - that line is only
                1.5-2.5px wide, far too thin to reliably click on its own.
                pointer-events is re-enabled here specifically; .edgeOverlay
                itself stays pointer-events:none so empty canvas space
                still click-through to the grid underneath it. The general
                sibling selector below (.edgeHitArea:hover ~ .edgeLine)
                relies on this exact order to highlight the visible line
                on hover, so it reads as "this arrow" rather than a vague
                hover with no visible feedback.
              */}
              {!isLocked && (
                <path
                  className={styles.edgeHitArea}
                  d={path}
                  stroke="transparent"
                  strokeWidth={14}
                  fill="none"
                  onClick={e => { e.stopPropagation(); setEdgePopup({ edge, point: { x: e.clientX, y: e.clientY } }); }}
                >
                  <title>Click to label or remove this dependency</title>
                </path>
              )}
              <path
                className={styles.edgeLine}
                d={path}
                stroke={highlighted ? '#1441b9' : '#3c4a63'}
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
        Anchored to the click's viewport point (see edgePopup), not a DOM
        element ref - there's no single stable element per edge the way
        there would be for an ordinary button. Both actions live in one
        popup rather than a menu you'd pick a branch from first, since
        seeing the current label (if any) while deciding whether to
        relabel or just remove it is more useful than hiding one behind
        the other.
      */}
      {edgePopup && (
        <Callout
          target={edgePopup.point}
          onDismiss={() => setEdgePopup(undefined)}
          setInitialFocus
          className={styles.edgePopup}
        >
          <TextField
            label="Branch label"
            placeholder="Yes / No / label this branch"
            defaultValue={edgePopup.edge.label || ''}
            onChange={(_e, v) => setDraftLabels(prev => ({ ...prev, [draftKey(edgePopup.edge)]: v || '' }))}
            onBlur={() => saveLabel(edgePopup.edge)}
          />
          <DefaultButton
            text="Remove dependency"
            iconProps={{ iconName: 'Delete', styles: { root: { color: 'var(--risk-high)' } } }}
            onClick={() => { deleteEdge(edgePopup.edge); setEdgePopup(undefined); }}
          />
        </Callout>
      )}

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
              const isDragOverThisCell = dragOverCellId === cellKey;
              const cellClassName = [
                styles.laneCell,
                isValidDropTargetForCurrentDrag(step) ? styles.validDropTarget : '',
                isDragOverThisCell ? styles.dragOver : '',
                isDragOverThisCell && !draggingNewShape
                  ? (dragOverPosition === 'before' ? styles.insertBefore : styles.insertAfter)
                  : '',
                // A grid full of bare, empty cell outlines reads as visual
                // noise in an exported PDF meant to be shared/read, even
                // though the same outlines are genuinely useful in the live
                // view (they're real drop targets there) - only hidden
                // during the export capture, never in normal editing.
                !belongsToLane && exporting ? styles.emptyCellExport : ''
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
                      draggable={!isLocked}
                      onDragStart={handleDragStart(step)}
                      onDragEnd={handleDragEnd}
                      onDragOver={handleConnectDragOver(step)}
                      onDrop={handleConnectDrop(step)}
                    >
                      <ShapeNode
                        label={step.actionDescription}
                        shape={getShapeType(step)}
                        linkedRiskSeverity={worstLinkedSeverity(step.linkedRisks)}
                        linkedRiskCount={(step.linkedRisks || []).length}
                        selected={selectedNodeId === step.id}
                        onClick={() => handleNodeClick(step)}
                        // Editing UI, not diagram content - hidden during
                        // PDF export the same way empty cells are (see
                        // cellClassName above), just via not rendering it
                        // at all rather than a CSS class, so there's no
                        // transition-timing risk to worry about either.
                        onConnectorDragStart={isLocked || exporting ? undefined : handleConnectDragStart(step)}
                        onConnectorDragEnd={handleConnectDragEnd}
                        connectable={isValidConnectionTarget(step)}
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
                <IconButton iconProps={{ iconName: 'Delete' }} title="Delete this task" onClick={deleteSelected} disabled={isLocked} />
                <IconButton iconProps={{ iconName: 'Cancel' }} title="Close" onClick={closeEditPopup} />
              </div>
            </div>

            {isLocked && <p className={styles.lockedNotice}>This swimlane is locked - viewing only.</p>}

            {(() => {
              const original = selectedNodeId ? stepsById.get(selectedNodeId) : undefined;
              if (!original || (!original.createdBy && !original.modifiedBy)) return null;
              return (
                <p className={styles.auditNotice}>
                  {original.createdBy && `Created by ${original.createdBy}${original.createdAt ? ` on ${new Date(original.createdAt).toLocaleString()}` : ''}`}
                  {original.createdBy && original.modifiedBy && ' · '}
                  {original.modifiedBy && `Last modified by ${original.modifiedBy}${original.modifiedAt ? ` on ${new Date(original.modifiedAt).toLocaleString()}` : ''}`}
                </p>
              );
            })()}

            <div className={styles.editForm}>
              <ProcessStepForm
                value={editDraft}
                onChange={setEditDraft}
                employees={employees}
                dependsOnOptions={dependsOnOptions}
                riskStatements={riskStatements}
              />
              <PrimaryButton text="Save changes" onClick={() => { saveEdit(); closeEditPopup(); }} disabled={isLocked} />
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
