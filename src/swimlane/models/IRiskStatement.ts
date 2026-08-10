/**
 * PLACEHOLDER MODEL. The "risk regnew" list's full column list was
 * explicitly flagged as not fully catalogued yet - only Title and Risk
 * Statement are confirmed. Everything else here is a reasonable guess at
 * what a risk register needs, not a confirmed schema - revisit once the
 * real list's internal field names are known.
 */
export interface IRiskStatement {
  id: string;
  title: string;
  riskStatement: string;
  linkedProcessStepIds?: string[]; // guess: how a risk might link back to specific process steps
}
