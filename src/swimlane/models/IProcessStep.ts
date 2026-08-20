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
  // RESPONSIBLE PERSON sits in, unrelated). A Process ID can hold several
  // genuinely different swimlanes side by side, one per region - see
  // FlowRegionTabs, which is the only thing that reads/sets this. '' or
  // undefined means the step isn't tagged to any particular region and
  // only shows up under "All".
  region?: string;
  shapeOverride?: string; // raw SharePoint value, e.g. 'Approval', 'Decision', 'Process Step'
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
  // (IProcessIdLock). In real (Graph) usage these are read straight from
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
  // Two separate reference links a step can carry - confirmed 2026-08-19.
  // sopLink points to the detailed SOP and/or system guidance for doing
  // the step (e.g. a NetSuite how-to page); delegationOfAuthorityLink
  // points to the DoA document governing who's allowed to take the
  // action. Deliberately kept as two distinct fields rather than one
  // generic "links" list - they're different documents with different
  // audiences, and a step might have either, both, or neither. Both
  // undefined/blank for the vast majority of steps, so the edit panel
  // hides each behind its own "+ Add link" affordance (see
  // OptionalLinkField) instead of always showing two empty boxes.
  sopLink?: string;
  delegationOfAuthorityLink?: string;
}

/**
 * Process ID groups a continuous flow of Process Step IDs (e.g. 9.6.1
 * contains 9.6.1.1 through 9.6.1.6 as one flow). There is no separate
 * SharePoint column for it - confirmed design rule is that it's always
 * derived by truncating Process Step ID to its first three dot-separated
 * segments.
 */
export function getProcessId(processStepId: string): string {
  return (processStepId || '').split('.').slice(0, 3).join('.');
}

// The regions always offered by FlowRegionTabs even before any step
// exists for them yet, and also what CSV import (see utils/csvImport.ts)
// recognizes when auto-detecting a region from an APQC Title's own
// trailing "- UK" suffix - shared here, in one place, rather than
// duplicated. 'Global' added 2026-08-19 for a process that genuinely
// applies everywhere rather than being tied to one country's variant -
// a real, taggable region in its own right, distinct from "All" (the
// FlowRegionTabs aggregate view showing every region's steps together
// regardless of tag, selectedFlowRegion === undefined).
export const KNOWN_FLOW_REGIONS = ['UK', 'US', 'SA', 'Global'];

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
 * Parses an Edge Labels cell into a token -> label map (a decision node's
 * Yes/No branch labels, set by clicking an edge on the canvas - see
 * handleLabelEdge in SwimlaneStudio.tsx). Same "token:value" pair format
 * as Linked Risks (see parseLinkedRisks in IRiskStatement.ts), comma/
 * newline separated - one pair per DependsOn token that has a label, e.g.
 * "9.6.1.1-3:Yes, 9.6.1.1-7:No". Unlike Linked Risks, a pair missing its
 * ":label" half is dropped rather than defaulted - there's no sensible
 * default for a branch label the way "Medium" is for a missing severity.
 */
export function parseEdgeLabels(raw: string | undefined): { [dependsOnToken: string]: string } {
  const result: { [dependsOnToken: string]: string } = {};
  (raw || '')
    .split(/[\n,]+/)
    .map(entry => entry.trim())
    .filter(entry => entry.length > 0)
    .forEach(entry => {
      const separatorIndex = entry.indexOf(':');
      if (separatorIndex <= 0) return;
      const token = entry.slice(0, separatorIndex).trim();
      const label = entry.slice(separatorIndex + 1).trim();
      if (token && label) result[token] = label;
    });
  return result;
}

export function serializeEdgeLabels(edgeLabels: { [dependsOnToken: string]: string } | undefined): string {
  return Object.entries(edgeLabels || {}).map(([token, label]) => `${token}:${label}`).join(', ');
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
