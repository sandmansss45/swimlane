import { IProcessStep } from '../models/IProcessStep';
import { dependsOnTokensToStepIds } from './dependencyResolution';

export interface IDuplicateRegionResult {
  steps: Array<Omit<IProcessStep, 'id'>>;
  // DependsOn tokens (and their matching edge label, if any) that pointed
  // at a step outside the set being duplicated - there's nothing sensible
  // to remap them to, since the step they referenced isn't coming along,
  // so they're dropped rather than left pointing at the wrong row.
  droppedDependencyCount: number;
}

/**
 * Builds the payload for "duplicate this region's steps into a new
 * region" (same Process ID, same Process Step IDs, everything else
 * copied) - used by DuplicateRegionModal. Reuses the same row-number
 * remapping approach as CSV import (see remapDependsOnToken in
 * csvImport.ts): `insertionIndex` is the current full dataset size, since
 * addProcessSteps guarantees a batch lands contiguously at the end (see
 * IDataService.addProcessSteps), and each duplicated step's position
 * within `sourceSteps` becomes its "local row number" (index + 2, header
 * counted as row 1) exactly like a row's position within an imported file.
 *
 * A DependsOn token pointing at a step that's part of `sourceSteps` gets
 * rewritten to point at that step's new row; one pointing outside it (a
 * cross-region dependency, or an already-unresolved token) is dropped -
 * see droppedDependencyCount. `allSteps` must be the full unfiltered
 * dataset, in original order, for token resolution to be correct (same
 * requirement as resolveDependencyEdges).
 */
export function buildDuplicatedSteps(
  allSteps: IProcessStep[],
  sourceSteps: IProcessStep[],
  targetRegion: string,
  insertionIndex: number
): IDuplicateRegionResult {
  const sourceIds = new Set(sourceSteps.map(s => s.id));
  const localRowNumberById = new Map<string, number>();
  sourceSteps.forEach((s, i) => localRowNumberById.set(s.id, i + 2));

  let droppedDependencyCount = 0;

  const steps = sourceSteps.map((step): Omit<IProcessStep, 'id'> => {
    // Resolved once per unique token on this step, then reused for both
    // dependsOn and edgeLabels below - edgeLabels keys are always a
    // subset of dependsOn tokens, so re-resolving them separately would
    // double-count the same dropped link.
    const tokenMap = new Map<string, string | undefined>();
    step.dependsOn.forEach(token => {
      if (tokenMap.has(token)) return;
      const [resolvedId] = dependsOnTokensToStepIds(allSteps, [token]);
      if (resolvedId && sourceIds.has(resolvedId)) {
        const localRow = localRowNumberById.get(resolvedId) as number;
        tokenMap.set(token, token.replace(/(\d+)\s*$/, () => String(insertionIndex + localRow)));
      } else {
        tokenMap.set(token, undefined);
        droppedDependencyCount += 1;
      }
    });

    const dependsOn = step.dependsOn
      .map(token => tokenMap.get(token))
      .filter((token): token is string => token !== undefined);

    const edgeLabels: { [token: string]: string } = {};
    Object.entries(step.edgeLabels || {}).forEach(([oldToken, label]) => {
      const newToken = tokenMap.get(oldToken);
      if (newToken) edgeLabels[newToken] = label;
    });

    return {
      apqcTitle: step.apqcTitle,
      processDescription: step.processDescription,
      processStepId: step.processStepId,
      processStepName: step.processStepName,
      actionType: step.actionType,
      action: step.action,
      actionDescription: step.actionDescription,
      responsibleJobTitle: step.responsibleJobTitle,
      region: targetRegion,
      shapeOverride: step.shapeOverride,
      manualOrder: step.manualOrder,
      dependsOn,
      edgeLabels: Object.keys(edgeLabels).length > 0 ? edgeLabels : undefined,
      linkedRisks: step.linkedRisks ? step.linkedRisks.map(link => ({ ...link })) : undefined,
      sopLink: step.sopLink,
      delegationOfAuthorityLink: step.delegationOfAuthorityLink
    };
  });

  return { steps, droppedDependencyCount };
}
