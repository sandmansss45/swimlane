import { IRiskLink } from './IRiskStatement';

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
  // Which country/entity's variant of this flow this step belongs to (e.g.
  // 'UK', 'US', 'SA') - confirmed 2026-08-18 as a real, separate concept
  // from IEmployee.department (which is just which department a
  // RESPONSIBLE PERSON sits in, unrelated). A Progress ID can hold several
  // genuinely different swimlanes side by side, one per region - see
  // FlowRegionTabs, which is the only thing that reads/sets this. '' or
  // undefined means the step isn't tagged to any particular region and
  // only shows up under "All".
  region?: string;
  shapeOverride?: string; // raw SharePoint value, e.g. 'Approval', 'Decision', 'Process Step'
  // Retired manual risk-level flag ('High' | 'Medium' | 'Low' | '') - no
  // longer settable from the UI and no longer drives any shape color (see
  // linkedRisks below, the app's one real risk indicator now). Field kept
  // only so existing values already stored in the real SharePoint column
  // still round-trip on load/save instead of silently vanishing.
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
  // Real Risk Register entries tied to this step, each with its own
  // manually-chosen severity (see IRiskLink) - drives a small corner
  // marker (see ShapeNode), the app's one real risk indicator.
  linkedRisks?: IRiskLink[];
  // Who created/last touched this step, and when - confirmed 2026-08-19,
  // part of the same audit-trail thread as swimlane locking
  // (IProgressIdLock). In real (Graph) usage these are read straight from
  // SharePoint's own native, system-managed item metadata (createdBy/
  // lastModifiedBy/createdDateTime/lastModifiedDateTime) - the app never
  // writes them as a custom column, since SharePoint already tracks this
  // accurately for every list item for free. In mock mode, and as an
  // optimistic same-session stamp in real mode (before the next reload
  // re-syncs with SharePoint's true values), the data services set these
  // themselves from whoever's actually signed in - see
  // MockDataService/GraphDataService constructors.
  createdBy?: string;
  createdAt?: string; // ISO date string
  modifiedBy?: string;
  modifiedAt?: string; // ISO date string
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

// The three regions always offered by FlowRegionTabs even before any step
// exists for them yet, and also what CSV import (see utils/csvImport.ts)
// recognizes when auto-detecting a region from an APQC Title's own
// trailing "- UK" suffix - shared here, in one place, rather than
// duplicated, so both only ever need updating once if a real 4th region
// shows up.
export const KNOWN_FLOW_REGIONS = ['UK', 'US', 'SA'];

/**
 * Real APQC Title values already carry their region as a trailing suffix
 * (e.g. "9.6.1 - UK") - confirmed against real data, predating the region
 * field itself. CSV import uses this so a file with that convention
 * already baked in lands its rows in the right swimlane without needing
 * a separate Region column at all. Deliberately restricted to
 * KNOWN_FLOW_REGIONS rather than matching any trailing "- XX" text, so an
 * unrelated abbreviation (e.g. a title genuinely ending "- AP") can't get
 * misread as a region.
 */
export function extractRegionFromApqcTitle(apqcTitle: string): string | undefined {
  const match = (apqcTitle || '').match(/-\s*([A-Za-z]+)\s*$/);
  if (!match) return undefined;
  const code = match[1].toUpperCase();
  return KNOWN_FLOW_REGIONS.includes(code) ? code : undefined;
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
