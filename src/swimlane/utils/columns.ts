import { IProcessStep } from '../models/IProcessStep';

/**
 * Tasks within a Progress ID group sequence left-to-right by numeric
 * Process Step ID order - confirmed design rule.
 */
export function compareProcessStepIds(a: string, b: string): number {
  const aParts = a.split('.').map(p => parseInt(p, 10));
  const bParts = b.split('.').map(p => parseInt(p, 10));
  const len = Math.max(aParts.length, bParts.length);
  for (let i = 0; i < len; i++) {
    const aVal = isNaN(aParts[i]) ? -1 : aParts[i];
    const bVal = isNaN(bParts[i]) ? -1 : bParts[i];
    if (aVal !== bVal) return aVal - bVal;
  }
  return 0;
}

/**
 * A step's position within its own Process Step ID group: its manualOrder
 * if it's ever been dragged, otherwise how many siblings (same Process
 * Step ID) precede it in the original source order - so a step that's
 * never been touched keeps behaving exactly like before (source order),
 * and mixing manually-moved and untouched siblings in the same group
 * still produces a sensible, stable result.
 */
export function getEffectiveOrder(step: IProcessStep, allSteps: IProcessStep[]): number {
  if (step.manualOrder !== undefined) return step.manualOrder;
  let index = 0;
  for (const s of allSteps) {
    if (s.id === step.id) break;
    if (s.processStepId === step.processStepId) index++;
  }
  return index;
}

export interface IColumnGroup {
  processStepId: string;
  stepIds: string[]; // in column order, left to right
}

/**
 * Every visible step gets its OWN column (a timeline slot), instead of
 * grouping every row that shares a Process Step ID into one narrow,
 * densely-stacked zone. A real swimlane reads left-to-right as one
 * continuous flow - cramming, say, six of one actor's steps into a single
 * cell forces every OTHER actor's edge into/out of that cell to detour
 * around the stack, which is exactly what was producing the tangled
 * gutter lines. Spreading every step across its own slot means most
 * connections are between near-adjacent cells with nothing else in the
 * way, which is what actually produces short, direct, uncluttered arrows.
 *
 * Sort is stable (JS guarantees this) and DependsOn row-number resolution
 * is keyed to the original unsorted dataset's own array index (see
 * resolveDependencyEdges), not to this function's output - so dragging a
 * step to a new column here never changes what its DependsOn tokens
 * resolve to, only where it's drawn.
 */
export function orderStepsForTimeline(steps: IProcessStep[]): IProcessStep[] {
  return [...steps].sort((a, b) => {
    const groupCompare = compareProcessStepIds(a.processStepId, b.processStepId);
    if (groupCompare !== 0) return groupCompare;
    return getEffectiveOrder(a, steps) - getEffectiveOrder(b, steps);
  });
}

/**
 * New manualOrder for dropping `draggedStepId` next to wherever
 * `targetStepId` currently sits within their shared Process Step ID group
 * - the fractional-midpoint drag-reorder technique (sits strictly between
 * the target's own position and whichever neighbour is on the requested
 * side), so only the dragged step's own record needs to change, not
 * everyone else's. `insertBefore` defaults to false (insert immediately
 * after the target, the original and still most common case) - true
 * inserts immediately before it instead. Without this, there was no way
 * to drop a step so it became the very FIRST item in its group: every
 * drop always landed after whatever cell it was dropped on, and there's
 * no column further left than the first one to drop "after" to get the
 * same effect (confirmed real user report - dragging to the front of a
 * group was simply impossible, not just awkward). See SwimlaneCanvas's
 * handleCellDragOver/handleDrop for how the two halves of a cell each
 * map to one of these.
 * Returns undefined if the two steps aren't actually in the same group -
 * dragging across groups is out of scope (see the manualOrder comment on
 * IProcessStep for why).
 */
export function computeDropOrder(
  allSteps: IProcessStep[],
  draggedStepId: string,
  targetStepId: string,
  insertBefore = false
): number | undefined {
  const dragged = allSteps.find(s => s.id === draggedStepId);
  const target = allSteps.find(s => s.id === targetStepId);
  if (!dragged || !target || dragged.processStepId !== target.processStepId || dragged.id === target.id) {
    return undefined;
  }
  const groupSteps = orderStepsForTimeline(
    allSteps.filter(s => s.processStepId === dragged.processStepId && s.id !== draggedStepId)
  );
  const targetIdx = groupSteps.findIndex(s => s.id === targetStepId);
  if (targetIdx === -1) return undefined;
  const targetOrder = getEffectiveOrder(groupSteps[targetIdx], allSteps);

  if (insertBefore) {
    const prevOrder = targetIdx - 1 >= 0
      ? getEffectiveOrder(groupSteps[targetIdx - 1], allSteps)
      : targetOrder - 1;
    return (prevOrder + targetOrder) / 2;
  }

  const nextOrder = targetIdx + 1 < groupSteps.length
    ? getEffectiveOrder(groupSteps[targetIdx + 1], allSteps)
    : targetOrder + 1;
  return (targetOrder + nextOrder) / 2;
}

/**
 * Same fractional-midpoint technique as computeDropOrder, for a step that
 * doesn't exist yet - dropping a shape from the legend onto an existing
 * column (see SwimlaneCanvas's onCreateStep) creates a brand new step
 * that takes over that column, shifting the column it was dropped onto
 * one place to the right - not appended after it, which would itself
 * render one column to the right of wherever the shape was actually
 * dropped (confirmed wrong against a real drop: the new step landed
 * "one square to the right" of the target instead of at it). No dragged
 * step to exclude from the group here, unlike computeDropOrder - that's
 * the only real difference between the two.
 */
export function computeInsertOrderBefore(allSteps: IProcessStep[], targetStepId: string): number | undefined {
  const target = allSteps.find(s => s.id === targetStepId);
  if (!target) return undefined;
  const groupSteps = orderStepsForTimeline(allSteps.filter(s => s.processStepId === target.processStepId));
  const targetIdx = groupSteps.findIndex(s => s.id === targetStepId);
  if (targetIdx === -1) return undefined;
  const targetOrder = getEffectiveOrder(groupSteps[targetIdx], allSteps);
  const prevOrder = targetIdx > 0
    ? getEffectiveOrder(groupSteps[targetIdx - 1], allSteps)
    : targetOrder - 1;
  return (prevOrder + targetOrder) / 2;
}

/**
 * Whether dropping `draggedStepId` next to `targetStepId` (before or
 * after, per `insertBefore` - i.e. at whatever order computeDropOrder
 * would give it) keeps every direct DependsOn relationship inside the
 * group pointing forward - confirmed design rule is that arrows must
 * keep moving in chronological order, so a step can't be dragged to sit
 * before something it depends on, or after something that depends on it.
 * Only edges where BOTH ends share the dragged step's Process Step ID
 * group can even be affected by an intra-group reorder - a cross-group
 * dependency's relative order never changes, since groups themselves
 * always stay in Process Step ID order regardless of manualOrder within
 * one of them.
 */
export function dropKeepsDependencyOrder(
  allSteps: IProcessStep[],
  edges: Array<{ fromRowId: string; toRowId: string }>,
  draggedStepId: string,
  targetStepId: string,
  insertBefore = false
): boolean {
  const dragged = allSteps.find(s => s.id === draggedStepId);
  const newOrder = computeDropOrder(allSteps, draggedStepId, targetStepId, insertBefore);
  if (!dragged || newOrder === undefined) return false;

  return edges.every(edge => {
    if (edge.fromRowId === draggedStepId) {
      const successor = allSteps.find(s => s.id === edge.toRowId);
      if (!successor || successor.processStepId !== dragged.processStepId) return true;
      return getEffectiveOrder(successor, allSteps) > newOrder;
    }
    if (edge.toRowId === draggedStepId) {
      const predecessor = allSteps.find(s => s.id === edge.fromRowId);
      if (!predecessor || predecessor.processStepId !== dragged.processStepId) return true;
      return getEffectiveOrder(predecessor, allSteps) < newOrder;
    }
    return true;
  });
}

/**
 * Groups consecutive same-Process-Step-ID columns for the header row,
 * so "9.6.1.1" still reads as one labeled zone spanning its own steps'
 * columns, the way a merged header cell would, without forcing every row
 * under it into the same physical cell.
 */
export function buildColumnGroups(orderedSteps: IProcessStep[]): IColumnGroup[] {
  const groups: IColumnGroup[] = [];
  orderedSteps.forEach(step => {
    const last = groups[groups.length - 1];
    if (last && last.processStepId === step.processStepId) {
      last.stepIds.push(step.id);
    } else {
      groups.push({ processStepId: step.processStepId, stepIds: [step.id] });
    }
  });
  return groups;
}
