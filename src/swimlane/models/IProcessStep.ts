export type ShapeType = 'process' | 'decision' | 'approval' | 'document';

export interface IProcessStep {
  id: string; // SharePoint list item ID
  apqcTitle: string;
  processDescription: string;
  processStepId: string; // e.g. '9.6.1.1'
  processStepName: string;
  actionType: string; // e.g. 'Execute (Within Limits)', 'Approve (Within Thresholds)', 'Endorse / Recommend'
  action: string;
  actionDescription: string;
  responsibleJobTitle: string;
  shapeOverride?: string; // raw SharePoint value, e.g. 'Approval', 'Decision', 'Process Step'
  // Raw value from the shape's edit panel ('High' | 'Medium' | 'Low' | '')
  // - kept as a plain string, same reasoning as shapeOverride, and
  // validated where it's consumed (resolveRiskLevel in IRiskStatement.ts)
  // rather than by the type here. Wins over whatever the Risk Register
  // would derive for this step; '' means "no override, use the Risk
  // Register".
  riskLevelOverride?: string;
  // Manual drag-and-drop position within this step's own Process Step ID
  // group (see orderStepsForTimeline in utils/columns.ts) - lower sorts
  // first. Independent of DependsOn on purpose: DependsOn only decides
  // which arrows connect to what, this decides left-to-right column
  // order. Undefined means "never manually moved", which falls back to
  // wherever it naturally landed in the source data.
  manualOrder?: number;
  // Raw DependsOn tokens, e.g. "9.6.1.1-3" - CONFIRMED these refer to a row
  // number (the row's position in the source data, header counted as row
  // 1 - so the first data row is row 2), not a Process Step ID. A single
  // cell can hold several tokens, newline- or comma-separated.
  dependsOn: string[];
  // Outgoing-edge labels (e.g. Yes/No) keyed by the raw DependsOn token
  // that produced the edge. Not part of the source data - there's nowhere
  // to import these from, so they're user-authored by clicking a decision
  // node. Every outgoing edge from a decision must have one; enforced in
  // the UI, not here.
  edgeLabels?: { [dependsOnToken: string]: string };
}

/**
 * Progress ID groups a continuous flow of Process Step IDs (e.g. 9.6.1
 * contains 9.6.1.1 through 9.6.1.6 as one flow). There is no separate
 * SharePoint column for it - confirmed design rule is that it's always
 * derived by truncating Process Step ID to its first three dot-separated
 * segments.
 */
export function getProgressId(processStepId: string): string {
  return (processStepId || '').split('.').slice(0, 3).join('.');
}

/**
 * Parses a DependsOn cell into individual raw tokens (e.g. "9.6.1.1-3").
 * The real export uses embedded newlines for multiple values within one
 * cell (confirmed against the actual CSV); comma is also accepted for
 * robustness.
 */
export function parseDependsOn(raw: string | undefined): string[] {
  return (raw || '')
    .split(/[\n,]+/)
    .map(token => token.trim())
    .filter(token => token.length > 0);
}

/**
 * Shapes: Approval-type = circles, Decision-type = diamonds, ordinary
 * process steps = rounded rectangles, physical/system artifacts
 * (purchase orders, invoices, remittances) = the document shape -
 * confirmed, do not flip this. ShapeOverride wins when present; Action
 * Type is the fallback signal when it's blank (Approve/Endorse actions
 * render as approval circles), and a question-phrased Action Description
 * is the last-resort signal for decisions the source data didn't
 * explicitly mark.
 */
export function getShapeType(step: IProcessStep): ShapeType {
  const override = (step.shapeOverride || '').trim().toLowerCase();
  if (override === 'approval') return 'approval';
  if (override === 'decision') return 'decision';
  if (override === 'process step') return 'process';
  if (override === 'document') return 'document';

  const actionType = (step.actionType || '').trim().toLowerCase();
  if (actionType.indexOf('approve') === 0 || actionType.indexOf('endorse') === 0) return 'approval';
  if (/\?\s*$/.test(step.actionDescription || '')) return 'decision';
  return 'process';
}

/**
 * Authority tier used to steer job-title suggestions: Execute -> analyst,
 * Endorse -> manager, Approve -> senior/chief. Confirmed design rule - do
 * not flip this mapping.
 */
export type AuthorityTier = 'analyst' | 'manager' | 'senior';

export function getAuthorityTier(actionType: string): AuthorityTier {
  const normalized = (actionType || '').trim().toLowerCase();
  if (normalized.indexOf('approve') === 0) return 'senior';
  if (normalized.indexOf('endorse') === 0) return 'manager';
  return 'analyst'; // Execute (Within Limits), Execute (Non Threshold), Automated, etc.
}
