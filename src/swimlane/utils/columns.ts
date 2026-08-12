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
 * Sort is stable (JS guarantees this), so rows sharing a Process Step ID
 * keep their original relative order - this doubles as the row-number
 * order DependsOn resolution already depends on, so it stays correct.
 */
export function orderStepsForTimeline(steps: IProcessStep[]): IProcessStep[] {
  return [...steps].sort((a, b) => compareProcessStepIds(a.processStepId, b.processStepId));
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
