import {
  isExcelErrorCode,
  parseA1,
  formatA1,
  rangeHeight,
  rangeWidth,
  type CellData,
  type CellFormat,
  type CellScalar,
  type ChangeOp,
  type RangeAddress,
  type WorkbookMap,
} from '@excelai/shared';
import { evaluateFormula, type CellResolver } from './formula.js';
import type { FindMatch, WorkbookExecutor } from './executor.js';

const FIND_LIMIT = 50;

interface EmuCell {
  raw: CellScalar;
  formula?: string;
}

/** Excel-like ascending order: numbers, then text, then booleans; blanks last. */
function compareSortKeys(a: CellScalar, b: CellScalar): number {
  const rank = (v: CellScalar): number =>
    v === null || v === '' ? 3 : typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2;
  const ra = rank(a);
  const rb = rank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  const sa = String(a).toLowerCase();
  const sb = String(b).toLowerCase();
  return sa < sb ? -1 : sa > sb ? 1 : 0;
}

/**
 * In-memory workbook with real formula recalculation. Implements the same
 * WorkbookExecutor interface as the Office.js bridge, so the entire agent
 * loop runs against it in tests and in the capability-probe sandbox.
 */
export class WorkbookEmulator implements WorkbookExecutor, CellResolver {
  private sheets = new Map<string, Map<string, EmuCell>>();
  private formats = new Map<string, Map<string, CellFormat>>();
  /** Synthetic stable worksheet ids, mirroring Office.js `worksheet.id`. */
  private sheetIds = new Map<string, string>();
  private nextSheetId = 1;
  /** cells currently being evaluated, for cycle detection */
  private evaluating = new Set<string>();

  constructor(sheetNames: string[] = ['Sheet1']) {
    for (const name of sheetNames) this.createSheet(name);
  }

  private createSheet(name: string): void {
    this.sheets.set(name, new Map());
    this.formats.set(name, new Map());
    this.sheetIds.set(name, `emu-sheet-${this.nextSheetId++}`);
  }

  /** Effective format of a single cell (test/inspection helper). */
  getFormat(ref: string): CellFormat | undefined {
    const addr = parseA1(ref, this.firstSheet());
    return this.formats.get(addr.sheet)?.get(`${addr.startRow},${addr.startCol}`);
  }

  /* ---------- CellResolver (formula engine callbacks) ---------- */

  sheetExists(sheet: string): boolean {
    return this.sheets.has(sheet);
  }

  getCell(sheet: string, row: number, col: number): CellScalar {
    const key = `${sheet}\0${row},${col}`;
    if (this.evaluating.has(key)) return '#CYCLE!';
    const cell = this.sheets.get(sheet)?.get(`${row},${col}`);
    if (!cell) return null;
    if (cell.formula !== undefined) {
      this.evaluating.add(key);
      try {
        return evaluateFormula(cell.formula, sheet, this);
      } finally {
        this.evaluating.delete(key);
      }
    }
    return cell.raw;
  }

  getRange(addr: RangeAddress): CellScalar[][] {
    const out: CellScalar[][] = [];
    for (let r = addr.startRow; r <= addr.endRow; r++) {
      const row: CellScalar[] = [];
      for (let c = addr.startCol; c <= addr.endCol; c++) {
        row.push(this.getCell(addr.sheet, r, c));
      }
      out.push(row);
    }
    return out;
  }

  /* ---------- direct manipulation (test setup) ---------- */

  setCell(ref: string, value: CellScalar | { formula: string }): void {
    const addr = parseA1(ref, this.firstSheet());
    const sheet = this.sheets.get(addr.sheet);
    if (!sheet) throw new Error(`No such sheet: ${addr.sheet}`);
    const key = `${addr.startRow},${addr.startCol}`;
    if (typeof value === 'object' && value !== null && 'formula' in value) {
      sheet.set(key, { raw: null, formula: value.formula });
    } else {
      sheet.set(key, { raw: value });
    }
  }

  getValue(ref: string): CellScalar {
    const addr = parseA1(ref, this.firstSheet());
    return this.getCell(addr.sheet, addr.startRow, addr.startCol);
  }

  private firstSheet(): string {
    const first = this.sheets.keys().next();
    if (first.done) throw new Error('Workbook has no sheets');
    return first.value;
  }

  /* ---------- WorkbookExecutor ---------- */

  async getWorkbookMap(): Promise<WorkbookMap> {
    const sheets = [...this.sheets.entries()].map(([name, cells]) => {
      let maxRow = -1;
      let maxCol = -1;
      let minRow = Number.MAX_SAFE_INTEGER;
      let minCol = Number.MAX_SAFE_INTEGER;
      for (const key of cells.keys()) {
        const [r, c] = key.split(',').map(Number) as [number, number];
        maxRow = Math.max(maxRow, r);
        maxCol = Math.max(maxCol, c);
        minRow = Math.min(minRow, r);
        minCol = Math.min(minCol, c);
      }
      const used =
        maxRow >= 0
          ? formatA1({ sheet: name, startRow: minRow, startCol: minCol, endRow: maxRow, endCol: maxCol })
          : null;
      return {
        id: this.sheetIds.get(name)!,
        name,
        rowCount: maxRow + 1,
        colCount: maxCol + 1,
        usedRange: used,
      };
    });
    return { sheets };
  }

  async readRange(range: string): Promise<CellData[][]> {
    const addr = parseA1(range, this.firstSheet());
    if (!this.sheets.has(addr.sheet)) throw new Error(`No such sheet: ${addr.sheet}`);
    const out: CellData[][] = [];
    for (let r = addr.startRow; r <= addr.endRow; r++) {
      const row: CellData[] = [];
      for (let c = addr.startCol; c <= addr.endCol; c++) {
        const cell = this.sheets.get(addr.sheet)!.get(`${r},${c}`);
        const value = this.getCell(addr.sheet, r, c);
        row.push({
          value,
          ...(cell?.formula !== undefined ? { formula: cell.formula } : {}),
          ...(isExcelErrorCode(value) ? { error: value as string } : {}),
        });
      }
      out.push(row);
    }
    return out;
  }

  async find(query: string, sheet?: string): Promise<FindMatch[]> {
    const q = query.toLowerCase();
    const matches: FindMatch[] = [];
    const names = sheet ? [sheet] : [...this.sheets.keys()];
    for (const name of names) {
      const cells = this.sheets.get(name);
      if (!cells) throw new Error(`No such sheet: ${sheet}`);
      const keys = [...cells.keys()]
        .map((k) => k.split(',').map(Number) as [number, number])
        .sort((a, b) => a[0] - b[0] || a[1] - b[1]);
      for (const [r, c] of keys) {
        const value = this.getCell(name, r, c);
        if (value !== null && String(value).toLowerCase().includes(q)) {
          matches.push({
            address: formatA1({ sheet: name, startRow: r, startCol: c, endRow: r, endCol: c }),
            value: String(value),
          });
          if (matches.length >= FIND_LIMIT) return matches;
        }
      }
    }
    return matches;
  }

  /** Rebuild a sheet's cell/format maps through a coordinate mapping. */
  private remap(
    sheetName: string,
    map: (row: number, col: number) => [number, number] | null,
  ): void {
    const cells = this.sheets.get(sheetName);
    const formats = this.formats.get(sheetName);
    if (!cells || !formats) throw new Error(`No such sheet: ${sheetName}`);
    const newCells = new Map<string, EmuCell>();
    const newFormats = new Map<string, CellFormat>();
    for (const [key, cell] of cells) {
      const [r, c] = key.split(',').map(Number) as [number, number];
      const target = map(r, c);
      if (target) newCells.set(`${target[0]},${target[1]}`, cell);
    }
    for (const [key, format] of formats) {
      const [r, c] = key.split(',').map(Number) as [number, number];
      const target = map(r, c);
      if (target) newFormats.set(`${target[0]},${target[1]}`, format);
    }
    this.sheets.set(sheetName, newCells);
    this.formats.set(sheetName, newFormats);
  }

  async applyOps(ops: ChangeOp[]): Promise<void> {
    for (const op of ops) {
      switch (op.kind) {
        case 'add_sheet': {
          if (this.sheets.has(op.name)) throw new Error(`Sheet already exists: ${op.name}`);
          this.createSheet(op.name);
          break;
        }
        case 'rename_sheet': {
          if (!this.sheets.has(op.name)) throw new Error(`No such sheet: ${op.name}`);
          if (this.sheets.has(op.newName)) throw new Error(`Sheet already exists: ${op.newName}`);
          this.sheets.set(op.newName, this.sheets.get(op.name)!);
          this.formats.set(op.newName, this.formats.get(op.name)!);
          this.sheetIds.set(op.newName, this.sheetIds.get(op.name)!);
          this.sheets.delete(op.name);
          this.formats.delete(op.name);
          this.sheetIds.delete(op.name);
          break;
        }
        case 'delete_sheet': {
          if (!this.sheets.has(op.name)) throw new Error(`No such sheet: ${op.name}`);
          if (this.sheets.size === 1) throw new Error('Cannot delete the only sheet in the workbook');
          this.sheets.delete(op.name);
          this.formats.delete(op.name);
          this.sheetIds.delete(op.name);
          break;
        }
        case 'insert_rows':
          this.remap(op.sheet, (r, c) => [r >= op.at ? r + op.count : r, c]);
          break;
        case 'delete_rows':
          this.remap(op.sheet, (r, c) =>
            r >= op.at && r < op.at + op.count ? null : [r >= op.at + op.count ? r - op.count : r, c],
          );
          break;
        case 'insert_cols':
          this.remap(op.sheet, (r, c) => [r, c >= op.at ? c + op.count : c]);
          break;
        case 'delete_cols':
          this.remap(op.sheet, (r, c) =>
            c >= op.at && c < op.at + op.count ? null : [r, c >= op.at + op.count ? c - op.count : c],
          );
          break;
        case 'sort_range': {
          const addr = parseA1(op.range, this.firstSheet());
          const cells = this.sheets.get(addr.sheet);
          const formats = this.formats.get(addr.sheet);
          if (!cells || !formats) throw new Error(`No such sheet: ${addr.sheet}`);
          const width = rangeWidth(addr);
          if (op.keyOffset < 0 || op.keyOffset >= width) {
            throw new Error('Sort key column is outside the range');
          }
          const firstDataRow = addr.startRow + (op.hasHeader ? 1 : 0);
          interface RowSnapshot {
            key: CellScalar;
            cells: (EmuCell | undefined)[];
            formats: (CellFormat | undefined)[];
          }
          const rows: RowSnapshot[] = [];
          for (let r = firstDataRow; r <= addr.endRow; r++) {
            const rowCells: (EmuCell | undefined)[] = [];
            const rowFormats: (CellFormat | undefined)[] = [];
            for (let c = addr.startCol; c <= addr.endCol; c++) {
              rowCells.push(cells.get(`${r},${c}`));
              rowFormats.push(formats.get(`${r},${c}`));
            }
            rows.push({
              key: this.getCell(addr.sheet, r, addr.startCol + op.keyOffset),
              cells: rowCells,
              formats: rowFormats,
            });
          }
          rows.sort((a, b) => {
            // Blanks stay last regardless of direction, like Excel.
            const blankA = a.key === null || a.key === '';
            const blankB = b.key === null || b.key === '';
            if (blankA || blankB) return Number(blankA) - Number(blankB);
            return compareSortKeys(a.key, b.key) * (op.ascending ? 1 : -1);
          });
          for (const [i, row] of rows.entries()) {
            const r = firstDataRow + i;
            for (let cIdx = 0; cIdx < width; cIdx++) {
              const key = `${r},${addr.startCol + cIdx}`;
              const cell = row.cells[cIdx];
              const format = row.formats[cIdx];
              if (cell) cells.set(key, cell);
              else cells.delete(key);
              if (format) formats.set(key, format);
              else formats.delete(key);
            }
          }
          break;
        }
        case 'format_range': {
          const addr = parseA1(op.range, this.firstSheet());
          const sheetFormats = this.formats.get(addr.sheet);
          if (!sheetFormats) throw new Error(`No such sheet: ${addr.sheet}`);
          const { autofitColumns: _autofit, ...cellFormat } = op.format;
          for (let r = addr.startRow; r <= addr.endRow; r++) {
            for (let c = addr.startCol; c <= addr.endCol; c++) {
              const key = `${r},${c}`;
              sheetFormats.set(key, { ...sheetFormats.get(key), ...cellFormat });
            }
          }
          break;
        }
        case 'clear_range': {
          const addr = parseA1(op.range, this.firstSheet());
          const sheet = this.sheets.get(addr.sheet);
          if (!sheet) throw new Error(`No such sheet: ${addr.sheet}`);
          for (let r = addr.startRow; r <= addr.endRow; r++) {
            for (let c = addr.startCol; c <= addr.endCol; c++) {
              sheet.delete(`${r},${c}`);
            }
          }
          break;
        }
        case 'write_range': {
          const addr = parseA1(op.range, this.firstSheet());
          const sheet = this.sheets.get(addr.sheet);
          if (!sheet) throw new Error(`No such sheet: ${addr.sheet}`);
          const h = rangeHeight(addr);
          const w = rangeWidth(addr);
          if (op.cells.length !== h || op.cells.some((row) => row.length !== w)) {
            throw new Error(
              `Cell matrix ${op.cells.length}x${op.cells[0]?.length ?? 0} does not match range ${op.range} (${h}x${w})`,
            );
          }
          for (let r = 0; r < h; r++) {
            for (let c = 0; c < w; c++) {
              const cell = op.cells[r]![c]!;
              const key = `${addr.startRow + r},${addr.startCol + c}`;
              if (cell.formula !== undefined) {
                sheet.set(key, { raw: null, formula: cell.formula });
              } else {
                sheet.set(key, { raw: cell.value ?? null });
              }
            }
          }
          break;
        }
      }
    }
  }
}
