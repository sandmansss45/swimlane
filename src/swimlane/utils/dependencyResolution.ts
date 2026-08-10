import { IProcessStep } from '../models/IProcessStep';

export interface IResolvedEdge {
  fromRowId: string;
  toRowId: string;
  token: string; // the raw DependsOn token that produced this edge, e.g. "9.6.1.1-3" - used as the edgeLabels key
  label?: string;
}

/**
 * CONFIRMED: DependsOn tokens (e.g. "9.6.1.1-3") reference a row by its
 * position in the source data, not a Process Step ID - the trailing
 * number is a spreadsheet row number with the header counted as row 1, so
 * the first data row is row 2, the second is row 3, and so on.
 *
 * IMPORTANT: `steps` must be in original source order (unfiltered) for
 * this numbering to be correct - row numbers are meaningless once rows
 * have been reordered or filtered out. Resolve edges once against the
 * full dataset immediately after loading, then let the rendering layer
 * decide which resolved edges have both endpoints currently visible.
 */
export function resolveDependencyEdges(orderedSteps: IProcessStep[]): IResolvedEdge[] {
  const rowNumberToIndex = new Map<number, number>();
  orderedSteps.forEach((_step, idx) => {
    rowNumberToIndex.set(idx + 2, idx); // idx 0 -> row 2 (header is row 1)
  });

  const edges: IResolvedEdge[] = [];
  orderedSteps.forEach((step, idx) => {
    step.dependsOn.forEach(token => {
      const match = /(\d+)\s*$/.exec(token);
      if (!match) return;
      const targetIdx = rowNumberToIndex.get(parseInt(match[1], 10));
      if (targetIdx === undefined || targetIdx === idx) {
        return; // unresolved token, or a row referencing itself
      }
      edges.push({
        fromRowId: orderedSteps[targetIdx].id,
        toRowId: step.id,
        token,
        label: step.edgeLabels && step.edgeLabels[token] ? step.edgeLabels[token] : undefined
      });
    });
  });
  return edges;
}

/**
 * Every outgoing edge from a decision must have a visible label (confirmed
 * design rule) - flags edges that need one.
 */
export function findUnlabeledDecisionEdges(steps: IProcessStep[], edges: IResolvedEdge[], getShapeType: (step: IProcessStep) => string): IResolvedEdge[] {
  const stepById = new Map(steps.map(s => [s.id, s]));
  return edges.filter(edge => {
    const fromStep = stepById.get(edge.fromRowId);
    if (!fromStep) return false;
    return getShapeType(fromStep) === 'decision' && !edge.label;
  });
}
