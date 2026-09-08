export type RiskLevel = 'High' | 'Medium' | 'Low';

/**
 * The real "risk register data" SharePoint list - confirmed columns as of
 * 2026-08-17 from a live screenshot of the list: Risk ID, Category, Risk
 * Statement, Root Cause, Risk Response. This is a standing enterprise
 * register this app doesn't own, but can add new rows to (see
 * IDataService.addRiskStatement - same deliberate, explicitly-confirmed
 * exception as addEmployee) - Risk ID is often blank in practice, so
 * linking (see IRiskLink below) keys off the SharePoint item id instead,
 * never the Risk ID text column.
 *
 * Likelihood (P), Materiality ($Mn), and Inherent Risk Rating were removed
 * 2026-09-08 - confirmed (live, on the real site) that these columns no
 * longer exist on the real list, so the app no longer reads, writes, or
 * displays them anywhere.
 *
 * TODO-CONFIRM: "Risk Owner" added 2026-09-08 at the user's request - not
 * yet verified against a live screenshot of the real list like the columns
 * above. If the real column has a different display name, get(item, 'Risk
 * Owner') in GraphDataService.getRiskStatements will silently come back
 * blank (console warning: "No field found for display name 'Risk Owner'")
 * rather than throwing - check the real list and update the display name
 * used there (and in addRiskStatement) once confirmed.
 */
export interface IRiskStatement {
  id: string; // SharePoint list item ID - the stable key used for linking
  riskId: string; // the "Risk ID" text column - often blank in the real data, display only
  category: string; // e.g. "External / Market / Reputational" - the grouping the risk-linking picker drills through
  riskStatement: string;
  rootCause: string;
  riskResponse: string; // e.g. "Mitigate", "Transfer"
  riskOwner: string; // TODO-CONFIRM - see interface comment above
}

/**
 * A risk tied to one specific process step. Severity is chosen by hand by
 * whoever ties it, not derived from the register's own numbers - confirmed
 * design rule: the same register risk can reasonably read as more or less
 * severe depending on which step it's attached to, and the person doing
 * the tying is better placed to judge that in the moment than a fixed
 * formula would be.
 */
export interface IRiskLink {
  riskId: string; // IRiskStatement.id, NOT the Risk ID text column
  severity: RiskLevel;
}

const SEVERITY_RANK: Record<RiskLevel, number> = { High: 3, Medium: 2, Low: 1 };

function toRiskLevel(raw: string | undefined): RiskLevel | undefined {
  const normalized = (raw || '').trim().toLowerCase();
  if (normalized === 'high') return 'High';
  if (normalized === 'medium') return 'Medium';
  if (normalized === 'low') return 'Low';
  return undefined;
}

/**
 * Raw column format is "riskId:severity" pairs, comma/newline separated
 * (e.g. "42:High, 51:Medium") - same delimiter convention as DependsOn
 * (see parseDependsOn in IProcessStep.ts). An entry with an unrecognized
 * or missing severity defaults to Medium rather than being dropped, so a
 * hand-edited cell missing the ":severity" part still shows up as linked
 * instead of silently disappearing.
 */
export function parseLinkedRisks(raw: string | undefined): IRiskLink[] {
  return (raw || '')
    .split(/[\n,]+/)
    .map(token => token.trim())
    .filter(token => token.length > 0)
    .map(token => {
      const [riskId, severityRaw] = token.split(':').map(part => (part || '').trim());
      return { riskId, severity: toRiskLevel(severityRaw) || 'Medium' };
    })
    .filter(link => link.riskId.length > 0);
}

export function serializeLinkedRisks(links: IRiskLink[]): string {
  return links.map(link => `${link.riskId}:${link.severity}`).join(', ');
}

/**
 * A step's shape gets a small corner marker when it has at least one
 * linked risk, colored by the WORST linked severity - "worst wins" is the
 * same reasoning the rest of this app's risk coloring already uses. This
 * is the app's one real risk indicator - a separate manual "risk level"
 * fill override used to exist alongside it (step.riskLevelOverride, set
 * directly from the edit panel without going through the Risk Register at
 * all) but was retired once this linked-register marker existed for real.
 */
export function worstLinkedSeverity(links: IRiskLink[] | undefined): RiskLevel | undefined {
  let worst: RiskLevel | undefined;
  (links || []).forEach(link => {
    if (!worst || SEVERITY_RANK[link.severity] > SEVERITY_RANK[worst]) worst = link.severity;
  });
  return worst;
}
