import { IProcessStep, parseDependsOn } from '../models/IProcessStep';

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
  'Action Type', 'Action', 'Action Description', 'ResponsibleJobTitle', 'ShapeOverride', 'DependsOn'
];

export interface IParsedCsvRow {
  step: Omit<IProcessStep, 'id'>;
  csvRowNumber: number; // position in the source file, header counted as row 1
  autoLinked: boolean; // true if DependsOn was blank and got a default chain
}

export interface ICsvImportPreview {
  rows: IParsedCsvRow[];
  warnings: string[];
  autoLinkedCount: number;
}

/**
 * Parses a CSV matching the confirmed export schema. Column order is not
 * assumed - each column is located by its display-name header, same
 * resolve-by-name approach the SharePoint data service uses, so the file
 * doesn't need to list columns in a fixed order.
 *
 * Auto-link heuristic: when a row's DependsOn is blank, it's chained to the
 * previous row within the same Process Step ID group (real exports list
 * steps in that group top-to-bottom in flow order, so this is the sensible
 * default). A row that's the first in its group with no DependsOn is left
 * with none - it's a legitimate starting point, not a missing value. This
 * is a deterministic fallback, not a trained model: there's no training
 * data or backend to run one on, and a default flow a person can correct
 * through the existing "Depends on" picker achieves the same outcome.
 */
export function buildImportPreview(csvText: string): ICsvImportPreview {
  const rawRows = parseCsvText(csvText);
  const warnings: string[] = [];

  if (rawRows.length === 0) {
    return { rows: [], warnings: ['The file is empty.'], autoLinkedCount: 0 };
  }

  const header = rawRows[0].map(h => h.trim());
  const headerIndex = new Map<string, number>();
  header.forEach((h, idx) => { if (h) headerIndex.set(h.toLowerCase(), idx); });

  const missing = REQUIRED_HEADERS.filter(h => !headerIndex.has(h.toLowerCase()));
  if (missing.length > 0) {
    warnings.push(`Missing required column(s): ${missing.join(', ')}. Check the file matches the expected export format.`);
    return { rows: [], warnings, autoLinkedCount: 0 };
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
  const lastRowNumberForGroup = new Map<string, number>();

  rawRows.forEach((cells, arrIdx) => {
    if (arrIdx === 0) return; // header
    const csvRowNumber = arrIdx + 1;
    if (cells.every(c => (c || '').trim() === '')) return; // blank separator row

    const processStepId = get(cells, 'Process Step ID');
    let dependsOn = parseDependsOn(get(cells, 'DependsOn'));
    let autoLinked = false;

    if (dependsOn.length === 0) {
      const priorRowNumber = lastRowNumberForGroup.get(processStepId);
      if (priorRowNumber !== undefined) {
        dependsOn = [`${processStepId || 'row'}-${priorRowNumber}`];
        autoLinked = true;
      }
    }

    if (processStepId) lastRowNumberForGroup.set(processStepId, csvRowNumber);

    parsedRows.push({
      csvRowNumber,
      autoLinked,
      step: {
        apqcTitle: get(cells, 'APQC Title'),
        processDescription: get(cells, 'Process Description'),
        processStepId,
        processStepName: get(cells, 'Process Step Name'),
        actionType: get(cells, 'Action Type'),
        action: get(cells, 'Action'),
        actionDescription: get(cells, 'Action Description'),
        responsibleJobTitle: get(cells, 'ResponsibleJobTitle'),
        shapeOverride: get(cells, 'ShapeOverride'),
        dependsOn
      }
    });
  });

  if (parsedRows.length === 0) {
    warnings.push('No data rows found.');
  }

  return { rows: parsedRows, warnings, autoLinkedCount: parsedRows.filter(r => r.autoLinked).length };
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
