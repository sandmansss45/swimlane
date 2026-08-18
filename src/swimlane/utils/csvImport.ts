import { IProcessStep, parseDependsOn, extractRegionFromApqcTitle } from '../models/IProcessStep';

/**
 * Minimal RFC4180 CSV parser (no external dependency): handles quoted
 * fields, embedded commas/newlines inside quotes, "" as an escaped quote,
 * and CRLF/LF line endings. The real export's DependsOn cells rely on
 * exactly this - a quoted field containing embedded newlines for multiple
 * tokens in one cell.
 */
export function parseCsvText(text: string): string[][] {
  const clean = text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;
  let i = 0;
  const n = clean.length;

  const pushField = (): void => { row.push(field); field = ''; };
  const pushRow = (): void => { pushField(); rows.push(row); row = []; };

  while (i < n) {
    const c = clean[i];
    if (inQuotes) {
      if (c === '"') {
        if (clean[i + 1] === '"') { field += '"'; i += 2; continue; }
        inQuotes = false; i += 1; continue;
      }
      field += c; i += 1; continue;
    }
    if (c === '"') { inQuotes = true; i += 1; continue; }
    if (c === ',') { pushField(); i += 1; continue; }
    if (c === '\r') { i += 1; continue; }
    if (c === '\n') { pushRow(); i += 1; continue; }
    field += c; i += 1;
  }
  if (field.length > 0 || row.length > 0) { pushRow(); }

  return rows.filter(r => !r.every(cell => (cell || '').trim() === ''));
}

const REQUIRED_HEADERS = ['Process Step ID', 'Process Step Name', 'Action Description'];
const KNOWN_HEADERS = [
  'APQC Title', 'Process Description', 'Process Step ID', 'Process Step Name',
  'Action Type', 'Action', 'Action Description', 'ResponsibleJobTitle', 'ShapeOverride', 'DependsOn', 'Region'
];

export interface IParsedCsvRow {
  step: Omit<IProcessStep, 'id'>;
  csvRowNumber: number; // position in the source file, header counted as row 1
  autoLinked: boolean; // true if DependsOn was blank and got a default chain
  autoRegioned: boolean; // true if Region was blank and got detected from the APQC Title's own "- UK" suffix
}

export interface ICsvImportPreview {
  rows: IParsedCsvRow[];
  warnings: string[];
  autoLinkedCount: number;
  autoRegionedCount: number;
}

/**
 * When a blank-DependsOn row's own Action verb has a known preferred
 * predecessor verb, prefer the most recent row in the group carrying that
 * verb over just grabbing whatever row happens to sit directly above it -
 * e.g. a "Receive" is really caused by a "Send", not by whatever
 * unrelated row a vendor happened to list first. Still plain keyword
 * matching against this export's own action vocabulary, not a trained
 * model reading the wording - it has no fallback for verbs outside this
 * list beyond "use the row before it", and no way to tell that a phrase
 * it wasn't specifically taught means the same thing.
 */
const VERB_PREDECESSORS: Record<string, string[]> = {
  receive: ['send', 'issue'],
  forward: ['send', 'receive'],
  review: ['receive', 'forward', 'create', 'submit'],
  approve: ['review', 'create', 'submit'],
  create: ['review', 'approve', 'receive'],
  submit: ['create', 'review'],
  automated: ['approve'],
  issue: ['approve', 'automated'],
  reconcile: ['issue', 'automated']
};

interface IGroupHistoryEntry {
  rowNumber: number;
  action: string;
}

/**
 * Picks the auto-link target for a blank-DependsOn row: the most recent
 * row in its group whose Action verb is a preferred predecessor for this
 * row's own verb, or - when this row's verb isn't in VERB_PREDECESSORS,
 * or none of its preferred verbs have appeared yet - the row directly
 * before it, same fallback as before this heuristic existed.
 */
function findAutoPredecessor(action: string, groupHistory: IGroupHistoryEntry[]): number | undefined {
  if (groupHistory.length === 0) return undefined;
  const preferred = VERB_PREDECESSORS[action.trim().toLowerCase()];
  if (preferred && preferred.length > 0) {
    for (let i = groupHistory.length - 1; i >= 0; i--) {
      if (preferred.includes(groupHistory[i].action.trim().toLowerCase())) {
        return groupHistory[i].rowNumber;
      }
    }
  }
  return groupHistory[groupHistory.length - 1].rowNumber;
}

/**
 * Parses a CSV matching the confirmed export schema. Column order is not
 * assumed - each column is located by its display-name header, same
 * resolve-by-name approach the SharePoint data service uses, so the file
 * doesn't need to list columns in a fixed order.
 *
 * Auto-link heuristic: when a row's DependsOn is blank, it's chained to a
 * row earlier in the same Process Step ID group - the row with a matching
 * verb per VERB_PREDECESSORS if one exists, otherwise the row directly
 * before it (real exports list steps in a group top-to-bottom in flow
 * order, so that's still a sensible default). A row that's the first in
 * its group with no DependsOn is left with none - it's a legitimate
 * starting point, not a missing value. This is a deterministic keyword
 * heuristic, not a trained model: there's no training data or backend to
 * run one on, and a default flow a person can correct through the
 * existing "Depends on" picker achieves the same outcome.
 */
export function buildImportPreview(csvText: string): ICsvImportPreview {
  const rawRows = parseCsvText(csvText);
  const warnings: string[] = [];

  if (rawRows.length === 0) {
    return { rows: [], warnings: ['The file is empty.'], autoLinkedCount: 0, autoRegionedCount: 0 };
  }

  const header = rawRows[0].map(h => h.trim());
  const headerIndex = new Map<string, number>();
  header.forEach((h, idx) => { if (h) headerIndex.set(h.toLowerCase(), idx); });

  const missing = REQUIRED_HEADERS.filter(h => !headerIndex.has(h.toLowerCase()));
  if (missing.length > 0) {
    warnings.push(`Missing required column(s): ${missing.join(', ')}. Check the file matches the expected export format.`);
    return { rows: [], warnings, autoLinkedCount: 0, autoRegionedCount: 0 };
  }

  const unknownColumns = header.filter(h => h && !KNOWN_HEADERS.some(k => k.toLowerCase() === h.toLowerCase()));
  if (unknownColumns.length > 0) {
    warnings.push(`Unrecognized column(s) ignored: ${unknownColumns.join(', ')}.`);
  }

  const get = (cells: string[], name: string): string => {
    const idx = headerIndex.get(name.toLowerCase());
    return idx === undefined ? '' : (cells[idx] || '').trim();
  };

  const parsedRows: IParsedCsvRow[] = [];
  const groupHistory = new Map<string, IGroupHistoryEntry[]>();
  // Row numbers skipped for missing a required value in that specific
  // cell, not just a missing column - REQUIRED_HEADERS above only checked
  // the header row HAS these columns, never that any given row's cells in
  // them weren't blank. An unnoticed blank Process Step ID used to import
  // fine and then silently vanish everywhere (it doesn't belong to any
  // Category/Group/Progress ID), while still consuming a row-number slot
  // other rows' DependsOn tokens count against.
  const skippedRows: Array<{ csvRowNumber: number; missingFields: string[] }> = [];

  rawRows.forEach((cells, arrIdx) => {
    if (arrIdx === 0) return; // header
    const csvRowNumber = arrIdx + 1;
    if (cells.every(c => (c || '').trim() === '')) return; // blank separator row

    const missingFields = REQUIRED_HEADERS.filter(h => !get(cells, h));
    if (missingFields.length > 0) {
      skippedRows.push({ csvRowNumber, missingFields });
      return;
    }

    const processStepId = get(cells, 'Process Step ID');
    const action = get(cells, 'Action');
    const apqcTitle = get(cells, 'APQC Title');
    let dependsOn = parseDependsOn(get(cells, 'DependsOn'));
    let autoLinked = false;

    if (dependsOn.length === 0) {
      const predecessorRow = findAutoPredecessor(action, groupHistory.get(processStepId) || []);
      if (predecessorRow !== undefined) {
        dependsOn = [`${processStepId || 'row'}-${predecessorRow}`];
        autoLinked = true;
      }
    }

    // An explicit Region column (if the file has one) always wins; failing
    // that, the real export's own naming convention already bakes the
    // region into APQC Title (e.g. "9.6.1 - UK") - detected here so a file
    // that already follows that convention lands its rows in the right
    // swimlane without needing a separate column added at all.
    let region = get(cells, 'Region');
    let autoRegioned = false;
    if (!region) {
      const detected = extractRegionFromApqcTitle(apqcTitle);
      if (detected) {
        region = detected;
        autoRegioned = true;
      }
    }

    if (processStepId) {
      const history = groupHistory.get(processStepId) || [];
      history.push({ rowNumber: csvRowNumber, action });
      groupHistory.set(processStepId, history);
    }

    parsedRows.push({
      csvRowNumber,
      autoLinked,
      autoRegioned,
      step: {
        apqcTitle,
        processDescription: get(cells, 'Process Description'),
        processStepId,
        processStepName: get(cells, 'Process Step Name'),
        actionType: get(cells, 'Action Type'),
        action,
        actionDescription: get(cells, 'Action Description'),
        responsibleJobTitle: get(cells, 'ResponsibleJobTitle'),
        shapeOverride: get(cells, 'ShapeOverride'),
        region,
        dependsOn
      }
    });
  });

  if (skippedRows.length > 0) {
    const rowList = skippedRows.map(s => `${s.csvRowNumber} (missing ${s.missingFields.join(', ')})`).join('; ');
    warnings.push(
      `Skipped ${skippedRows.length} row${skippedRows.length === 1 ? '' : 's'} with a blank required value - ` +
      `not imported: ${rowList}. Fix these in the file and re-upload if they should be included.`
    );
  }

  if (parsedRows.length === 0) {
    warnings.push('No data rows found.');
  }

  return {
    rows: parsedRows,
    warnings,
    autoLinkedCount: parsedRows.filter(r => r.autoLinked).length,
    autoRegionedCount: parsedRows.filter(r => r.autoRegioned).length
  };
}

/**
 * Rewrites a DependsOn token's trailing row number from "position within
 * the imported file" to "position within the live dataset it's landing
 * in". CONFIRMED math: a row at file row number R, inserted so the batch
 * starts at array index `insertionIndex` (0-based, i.e. the current step
 * count before import), ends up at the live row number
 * `insertionIndex + R` under the app's header-is-row-1 convention - see
 * resolveDependencyEdges. Only the trailing digits carry meaning; the
 * token's own prefix text is cosmetic and left as-is.
 */
export function remapDependsOnToken(token: string, insertionIndex: number): string {
  return token.replace(/(\d+)\s*$/, (_match, digits: string) => String(insertionIndex + parseInt(digits, 10)));
}

export function prepareStepsForImport(rows: IParsedCsvRow[], insertionIndex: number): Array<Omit<IProcessStep, 'id'>> {
  return rows.map(r => ({
    ...r.step,
    dependsOn: r.step.dependsOn.map(token => remapDependsOnToken(token, insertionIndex))
  }));
}
