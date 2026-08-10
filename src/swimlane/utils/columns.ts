import { IProcessStep } from '../models/IProcessStep';

/**
 * Tasks within a Progress ID group sequence left-to-right by numeric
 * Process Step ID order - confirmed design rule.
 *
 * Top-level view (all rows under a Progress ID): one column per distinct
 * Process Step ID, ordered numerically, so rows sharing a step stack
 * together in one column.
 *
 * Drill-down view (single Process Step ID selected): every row shares that
 * one step, so there's only one column - rows stack vertically within
 * their lane cell in original row order instead of spreading sideways
 * (spreading them into one column per row made a 10-row step 10 columns
 * wide for no reason, forcing pointless horizontal scrolling).
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

export function buildColumns(steps: IProcessStep[], drilledDownStepId: string | undefined): string[] {
  if (drilledDownStepId) {
    return [drilledDownStepId];
  }
  const distinctStepIds = Array.from(new Set(steps.map(s => s.processStepId)));
  return distinctStepIds.sort(compareProcessStepIds);
}

export function columnKeyFor(step: IProcessStep, _drilledDownStepId: string | undefined): string {
  return step.processStepId;
}
