import { IProcessStep } from '../models/IProcessStep';

// Same column set, same order, as documented in ImportCsvModal's own
// instructions text - so a file this produces is immediately valid input
// to Import CSV again, round-tripping through the exact same schema
// rather than a lookalike one.
const EXPORT_HEADERS = [
  'APQC Title', 'Process Description', 'Process Step ID', 'Process Step Name',
  'Action Type', 'Action', 'Action Description', 'ResponsibleJobTitle', 'ShapeOverride', 'DependsOn', 'Region'
];

/**
 * RFC4180 field escaping (the inverse of parseCsvText in csvImport.ts) -
 * quotes a field only when it actually needs it (contains a comma,
 * double-quote, or newline), doubling any embedded double-quotes.
 */
function escapeCsvField(value: string): string {
  if (!/[",\r\n]/.test(value)) return value;
  return `"${value.replace(/"/g, '""')}"`;
}

/**
 * Builds a CSV matching the confirmed import schema from the current live
 * dataset. Always the FULL steps array, in its current order, never a
 * filtered subset - DependsOn tokens are row numbers relative to that
 * exact order (header counted as row 1, see parseDependsOn), so exporting
 * anything less than the whole thing in its current order would silently
 * make some tokens point at the wrong row (or no row) once re-imported.
 * Multiple DependsOn tokens in one cell are newline-joined, matching the
 * same convention the real export and Import CSV already use.
 */
export function buildExportCsv(steps: IProcessStep[]): string {
  const lines = [EXPORT_HEADERS.join(',')];
  steps.forEach(step => {
    const row = [
      step.apqcTitle,
      step.processDescription,
      step.processStepId,
      step.processStepName,
      step.actionType,
      step.action,
      step.actionDescription,
      step.responsibleJobTitle,
      step.shapeOverride || '',
      step.dependsOn.join('\n'),
      step.region || ''
    ];
    lines.push(row.map(escapeCsvField).join(','));
  });
  // CRLF, not just LF - Import CSV's own parser accepts either, but CRLF
  // is what Excel (the actual destination for "offline review") expects
  // as a plain, unsurprising CSV rather than a Unix-style one.
  return lines.join('\r\n');
}

/**
 * Triggers a browser download of the given text as a file - no server
 * round-trip, just a Blob + a synthetic click on an object URL, same
 * technique as the existing PDF export in SwimlaneCanvas.tsx uses for its
 * own download. Leading BOM so Excel opens it as UTF-8 rather than
 * mis-detecting ANSI and mangling any non-ASCII text - the import side
 * already strips one for exactly this reason (see parseCsvText), meaning
 * real-world files in this app's ecosystem are already expected to carry
 * one.
 */
const UTF8_BOM = String.fromCharCode(0xfeff);

export function downloadTextFile(filename: string, content: string, mimeType: string): void {
  const blob = new Blob([UTF8_BOM + content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}
