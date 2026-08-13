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
 * New manualOrder for dropping `draggedStepId` immediately after wherever
 * `targetStepId` currently sits within their shared Process Step ID group
 * - the fractional-midpoint drag-reorder technique (sits strictly between
 * the target's own position and whatever comes after it), so only the
 * dragged step's own record needs to change, not everyone else's.
 * Returns undefined if the two steps aren't actually in the same group -
 * dragging across groups is out of scope (see the manualOrder comment on
 * IProcessStep for why).
 */
export function computeDropOrder(allSteps: IProcessStep[], draggedStepId: string, targetStepId: string): number | undefined {
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
  const nextOrder = targetIdx + 1 < groupSteps.length
    ? getEffectiveOrder(groupSteps[targetIdx + 1], allSteps)
    : targetOrder + 1;
  return (targetOrder + nextOrder) / 2;
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
