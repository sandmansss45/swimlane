import { IProcessStep } from './IProcessStep';

/**
 * PLACEHOLDER MODEL. The "risk regnew" list's full column list was
 * explicitly flagged as not fully catalogued yet - only Title and Risk
 * Statement are confirmed. Everything else here is a reasonable guess at
 * what a risk register needs, not a confirmed schema - revisit once the
 * real list's internal field names are known.
 */
export type RiskLevel = 'High' | 'Medium' | 'Low';

export interface IRiskStatement {
  id: string;
  title: string;
  riskStatement: string;
  linkedProcessStepIds?: string[]; // guess: how a risk might link back to specific process steps
  riskLevel?: RiskLevel; // guess: traffic-light severity, drives the step shape's fill color once linked
}

const SEVERITY_RANK: Record<RiskLevel, number> = { High: 3, Medium: 2, Low: 1 };

/**
 * A process step's shape is colored by the WORST linked risk, not just
 * "any" risk - if a step carries both a Medium and a High risk, showing
 * amber instead of red would understate what someone looking at the
 * diagram needs to notice first.
 */
export function getRiskLevel(processStepId: string, riskStatements: IRiskStatement[]): RiskLevel | undefined {
  let worst: RiskLevel | undefined;
  riskStatements.forEach(risk => {
    if (!risk.riskLevel) return;
    if (!(risk.linkedProcessStepIds || []).includes(processStepId)) return;
    if (!worst || SEVERITY_RANK[risk.riskLevel] > SEVERITY_RANK[worst]) worst = risk.riskLevel;
  });
  return worst;
}

function toRiskLevel(raw: string | undefined): RiskLevel | undefined {
  return raw === 'High' || raw === 'Medium' || raw === 'Low' ? raw : undefined;
}

/**
 * The level a step's shape should actually render with - a manual pick
 * from the edit panel (step.riskLevelOverride) wins over whatever the
 * Risk Register would derive, so someone can flag or clear a risk
 * directly on a step without having to go author/edit a Risk Register
 * entry first.
 */
export function resolveRiskLevel(step: IProcessStep, riskStatements: IRiskStatement[]): RiskLevel | undefined {
  const override = toRiskLevel(step.riskLevelOverride);
  if (override) return override;
  return getRiskLevel(step.processStepId, riskStatements);
}
