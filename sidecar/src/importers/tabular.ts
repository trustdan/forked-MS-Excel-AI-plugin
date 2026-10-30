import { existsSync } from 'node:fs';
import { extname } from 'node:path';
import {
  colToIndex,
  formatA1,
  indexToCol,
  parseA1,
  type CellScalar,
  type ChangeOp,
} from '@excelai/shared';
import { readFileCapped } from '../boundedRead.js';
import { PathGuard } from '../mcp/pathGuard.js';

export const MAX_IMPORT_BYTES = 10 * 1024 * 1024; // 10MB
export const MAX_IMPORT_CELLS = 50_000;
export const ALLOWED_EXTENSIONS = new Set(['.csv', '.tsv', '.txt', '.md', '.markdown', '.json']);

/**
 * Escapes leading formula characters to neutralize CSV / Formula Injection (CWE-1236).
 * Prepends a single quote so spreadsheet engines interpret it strictly as literal text.
 */
export function sanitizeFormulaInjection(val: string): string {
  if (!val) return val;
  const first = val[0];
  if (first === '=' || first === '@' || first === '+' || first === '-' || first === '\t' || first === '\r') {
    // If it's a valid integer or floating point number (e.g. -42, +3.14), it's safe as a number
    if (/^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$/.test(val)) {
      return val;
    }
    return `'${val}`;
  }
  return val;
}

/**
 * Infer the best JS scalar type for a cell string.
 */
export function inferCellValue(raw: string, sanitize = true): CellScalar {
  const trimmed = raw.trim();
  if (trimmed === '') return null;

  // Booleans
  const lower = trimmed.toLowerCase();
  if (lower === 'true') return true;
  if (lower === 'false') return false;

  // Numbers (strict integer, float, scientific notation)
  if (/^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$/.test(trimmed)) {
    const num = Number(trimmed);
    if (!Number.isNaN(num) && Number.isFinite(num)) {
      return num;
    }
  }

  // Currency / percentages
  const currencyMatch = /^\$([+-]?\d{1,3}(,\d{3})*(\.\d+)?)$/.exec(trimmed);
  if (currencyMatch && currencyMatch[1]) {
    const num = Number(currencyMatch[1].replace(/,/g, ''));
    if (!Number.isNaN(num)) return num;
  }

  const percentMatch = /^([+-]?\d+(\.\d+)?)\s*%$/.exec(trimmed);
  if (percentMatch && percentMatch[1]) {
    const num = Number(percentMatch[1]) / 100;
    if (!Number.isNaN(num)) return num;
  }

  return sanitize ? sanitizeFormulaInjection(trimmed) : trimmed;
}

/**
 * RFC 4180-compliant CSV/TSV parser supporting quotes, escaped quotes, multiline values.
 */
export function parseDelimited(text: string, delimiter: string = ','): string[][] {
  const rows: string[][] = [];
  let currentRow: string[] = [];
  let currentField = '';
  let inQuotes = false;

  const len = text.length;
  for (let i = 0; i < len; i++) {
    const char = text[i];
    const nextChar = text[i + 1];

    if (inQuotes) {
      if (char === '"') {
        if (nextChar === '"') {
          // Escaped quote
          currentField += '"';
          i++;
        } else {
          // End of quote
          inQuotes = false;
        }
      } else {
        currentField += char;
      }
    } else {
      if (char === '"') {
        inQuotes = true;
      } else if (char === delimiter) {
        currentRow.push(currentField);
        currentField = '';
      } else if (char === '\r') {
        if (nextChar === '\n') i++; // Skip CRLF
        currentRow.push(currentField);
        rows.push(currentRow);
        currentRow = [];
        currentField = '';
      } else if (char === '\n') {
        currentRow.push(currentField);
        rows.push(currentRow);
        currentRow = [];
        currentField = '';
      } else {
        currentField += char;
      }
    }
  }

  // Push final field/row if any data exists
  if (currentField.length > 0 || currentRow.length > 0) {
    currentRow.push(currentField);
    rows.push(currentRow);
  }

  // Filter out trailing blank row if input ended with newline
  const lastRow = rows[rows.length - 1];
  if (lastRow && lastRow.length === 1 && lastRow[0] === '') {
    rows.pop();
  }

  return rows;
}

/**
 * Parse GitHub-flavored Markdown tables.
 */
export function parseMarkdownTable(text: string): string[][] {
  const lines = text.split(/\r?\n/).map((l) => l.trim()).filter((l) => l.length > 0);
  const rows: string[][] = [];

  for (const line of lines) {
    if (!line.includes('|')) continue;

    // Check if delimiter row, e.g. |---|:---:|---:|
    if (/^\|?\s*[:\-]+(?:\s*\|\s*[:\-]+)+\s*\|?$/.test(line)) {
      continue;
    }

    let trimmed = line;
    if (trimmed.startsWith('|')) trimmed = trimmed.slice(1);
    if (trimmed.endsWith('|')) trimmed = trimmed.slice(0, -1);

    const cells = trimmed.split('|').map((c) => c.trim());
    if (cells.length > 0) {
      rows.push(cells);
    }
  }

  return rows;
}

/**
 * Parse JSON array-of-objects or array-of-arrays.
 */
export function parseJsonTable(text: string): string[][] {
  const parsed = JSON.parse(text);
  if (!Array.isArray(parsed)) {
    throw new Error('JSON data must be an array of objects or an array of arrays');
  }
  if (parsed.length === 0) return [];

  // Array of arrays
  if (Array.isArray(parsed[0])) {
    return parsed.map((row: unknown[]) => row.map((cell) => (cell === null || cell === undefined ? '' : String(cell))));
  }

  // Array of objects
  if (typeof parsed[0] === 'object' && parsed[0] !== null) {
    const keys = Array.from(
      new Set(parsed.flatMap((item: Record<string, unknown>) => Object.keys(item || {}))),
    );
    const rows: string[][] = [keys];
    for (const item of parsed) {
      const obj = (item || {}) as Record<string, unknown>;
      rows.push(keys.map((k) => (obj[k] === null || obj[k] === undefined ? '' : String(obj[k]))));
    }
    return rows;
  }

  throw new Error('Unsupported JSON shape for table parsing');
}

/**
 * Detect delimiter from text sample.
 */
export function detectDelimiter(sample: string): string {
  const lines = sample.split(/\r?\n/).slice(0, 5).filter((l) => l.trim().length > 0);
  if (lines.length === 0) return ',';

  let commaCount = 0;
  let tabCount = 0;
  let semiCount = 0;

  for (const line of lines) {
    commaCount += (line.match(/,/g) || []).length;
    tabCount += (line.match(/\t/g) || []).length;
    semiCount += (line.match(/;/g) || []).length;
  }

  if (tabCount > commaCount && tabCount > semiCount) return '\t';
  if (semiCount > commaCount && semiCount > tabCount) return ';';
  return ',';
}

/**
 * Parse raw tabular string with auto-detection.
 */
export function parseTabularContent(
  content: string,
  hintFormat?: 'csv' | 'tsv' | 'markdown' | 'json' | 'auto',
): CellScalar[][] {
  const trimmed = content.trim();
  if (!trimmed) return [];

  let rawRows: string[][];

  if (hintFormat === 'json' || (hintFormat === 'auto' && (trimmed.startsWith('[') || trimmed.startsWith('{')))) {
    try {
      rawRows = parseJsonTable(trimmed);
    } catch {
      rawRows = parseDelimited(trimmed, detectDelimiter(trimmed));
    }
  } else if (hintFormat === 'markdown' || (hintFormat === 'auto' && trimmed.includes('|') && /\|.*\|/.test(trimmed))) {
    const mdRows = parseMarkdownTable(trimmed);
    rawRows = mdRows.length > 0 ? mdRows : parseDelimited(trimmed, detectDelimiter(trimmed));
  } else if (hintFormat === 'tsv') {
    rawRows = parseDelimited(trimmed, '\t');
  } else {
    const delimiter = detectDelimiter(trimmed);
    rawRows = parseDelimited(trimmed, delimiter);
  }

  // Normalize column count to max row length
  const maxCols = Math.max(...rawRows.map((r) => r.length), 0);
  if (maxCols === 0) return [];

  return rawRows.map((row) => {
    const padded = [...row];
    while (padded.length < maxCols) padded.push('');
    return padded.map((cell) => inferCellValue(cell, true));
  });
}

/**
 * Read and parse local file with strict security and memory bounds. Only paths
 * inside `approvedPaths` (the user's approved folders) are ever opened.
 */
export function loadAndParseLocalFile(
  filePath: string,
  approvedPaths: readonly string[],
  options?: { maxBytes?: number; maxCells?: number },
): { rows: CellScalar[][]; format: string; totalRows: number; totalCols: number } {
  const maxBytes = options?.maxBytes ?? MAX_IMPORT_BYTES;
  const maxCells = options?.maxCells ?? MAX_IMPORT_CELLS;

  const resolved = PathGuard.assertApproved(filePath, [...approvedPaths], 'import');
  if (!existsSync(resolved)) {
    throw new Error(`File not found: "${filePath}"`);
  }

  const ext = extname(resolved).toLowerCase();
  if (!ALLOWED_EXTENSIONS.has(ext)) {
    throw new Error(
      `Unsupported file format "${ext}". Supported extensions: ${Array.from(ALLOWED_EXTENSIONS).join(', ')}`,
    );
  }

  const { bytes, size, truncated } = readFileCapped(resolved, maxBytes);
  if (truncated) {
    throw new Error(`File size (${Math.round(size / 1024)} KB) exceeds limit of ${Math.round(maxBytes / 1024)} KB.`);
  }

  let format: 'csv' | 'tsv' | 'markdown' | 'json' | 'auto' = 'auto';
  if (ext === '.tsv') format = 'tsv';
  else if (ext === '.json') format = 'json';
  else if (ext === '.md' || ext === '.markdown') format = 'markdown';
  else if (ext === '.csv') format = 'csv';

  const content = bytes.toString('utf-8');
  const rows = parseTabularContent(content, format);

  const totalRows = rows.length;
  const totalCols = rows[0]?.length ?? 0;
  const totalCells = totalRows * totalCols;

  if (totalCells > maxCells) {
    throw new Error(
      `Table dimensions (${totalRows} rows × ${totalCols} columns = ${totalCells} cells) exceed safety limit of ${maxCells} cells.`,
    );
  }

  return { rows, format, totalRows, totalCols };
}

/**
 * Calculate the target range address and generate staged change ops for importing tabular data.
 */
export function generateImportChangeOps(args: {
  rows: CellScalar[][];
  targetSheet: string;
  startCell?: string;
  hasHeaders?: boolean;
  reason: string;
}): ChangeOp[] {
  const { rows, targetSheet, startCell = 'A1', hasHeaders = true, reason } = args;
  const firstRow = rows[0];
  if (!firstRow || firstRow.length === 0) return [];

  // Parse startCell (e.g. "A1" or "Sheet1!B2")
  let startCol = 0;
  let startRow = 0;
  try {
    const parsed = parseA1(startCell, targetSheet);
    startCol = parsed.startCol;
    startRow = parsed.startRow;
  } catch {
    // Default to A1 if startCell is invalid
    startCol = 0;
    startRow = 0;
  }

  const endCol = startCol + firstRow.length - 1;
  const endRow = startRow + rows.length - 1;

  const targetRange = formatA1({
    sheet: targetSheet,
    startCol,
    startRow,
    endCol,
    endRow,
  });

  const writeCells = rows.map((r) => r.map((val) => ({ value: val })));

  const ops: ChangeOp[] = [
    {
      kind: 'write_range',
      range: targetRange,
      cells: writeCells,
      reason,
    },
  ];

  // If table has headers, add styling (bold + bottom border)
  if (hasHeaders && rows.length > 1) {
    const headerRange = formatA1({
      sheet: targetSheet,
      startCol,
      startRow,
      endCol,
      endRow: startRow,
    });
    ops.push({
      kind: 'format_range',
      range: headerRange,
      format: {
        bold: true,
        border: {
          edges: ['bottom'],
          style: 'medium',
        },
      },
      reason: `Format header row for imported data`,
    });
  }

  return ops;
}
