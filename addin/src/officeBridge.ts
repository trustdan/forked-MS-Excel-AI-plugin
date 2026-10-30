import {
  indexToCol,
  parseA1,
  type BorderFormat,
  type CellData,
  type ChangeOp,
  type WorkbookMap,
} from '@excelai/shared';

/**
 * The "hands": executes workbook tools against real Excel via Office.js.
 * Mirrors the sidecar's WorkbookEmulator behavior — the contract both
 * implementations share is defined by the shared tool schemas.
 */

export async function getWorkbookMap(): Promise<WorkbookMap> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load('items/name,items/id');
    await context.sync();

    const usedRanges = sheets.items.map((ws) => {
      const used = ws.getUsedRangeOrNullObject();
      used.load(['address', 'rowCount', 'columnCount', 'isNullObject']);
      return { id: ws.id, name: ws.name, used };
    });
    await context.sync();

    return {
      sheets: usedRanges.map(({ id, name, used }) => ({
        id,
        name,
        rowCount: used.isNullObject ? 0 : used.rowCount,
        colCount: used.isNullObject ? 0 : used.columnCount,
        usedRange: used.isNullObject ? null : used.address,
      })),
    };
  });
}

export async function readRange(rangeRef: string): Promise<CellData[][]> {
  return Excel.run(async (context) => {
    const resolver = await loadWorksheetResolver(context);
    const range = getRange(rangeRef, resolver);
    range.load(['values', 'formulas']);
    await context.sync();

    const out: CellData[][] = [];
    for (let r = 0; r < range.values.length; r++) {
      const row: CellData[] = [];
      for (let c = 0; c < range.values[r]!.length; c++) {
        const value = range.values[r]![c] as CellData['value'];
        const formula = range.formulas[r]![c];
        const isFormula = typeof formula === 'string' && formula.startsWith('=');
        row.push({ value, ...(isFormula ? { formula } : {}) });
      }
      out.push(row);
    }
    return out;
  });
}

const BATCH_SIZE = 50;

/**
 * Checks whether an op alters workbook topology or shifts row/column coordinates,
 * requiring a synchronization boundary before subsequent operations run.
 */
function isStructuralOp(op: ChangeOp): boolean {
  switch (op.kind) {
    case 'add_sheet':
    case 'rename_sheet':
    case 'delete_sheet':
    case 'insert_rows':
    case 'delete_rows':
    case 'insert_cols':
    case 'delete_cols':
      return true;
    default:
      return false;
  }
}

export async function applyOps(ops: ChangeOp[]): Promise<void> {
  if (ops.length === 0) return;

  await Excel.run(async (context) => {
    const resolver = await loadWorksheetResolver(context);
    let pendingOpsCount = 0;

    for (const op of ops) {
      // If the upcoming op is structural and we have pending queued mutations,
      // flush them first so the sheet is in a consistent state.
      if (isStructuralOp(op) && pendingOpsCount > 0) {
        await context.sync();
        pendingOpsCount = 0;
      }

      switch (op.kind) {
        case 'add_sheet': {
          const existing = resolver.findExisting(op.name);
          if (existing) throw new Error(`Sheet already exists: ${existing.name}`);
          const ws = context.workbook.worksheets.add(op.name);
          resolver.registerAdded(op.name, ws);
          // Structural: sync so the sheet exists in the workbook collection
          await context.sync();
          pendingOpsCount = 0;
          break;
        }
        case 'clear_range': {
          getRange(op.range, resolver).clear(Excel.ClearApplyTo.contents);
          pendingOpsCount++;
          break;
        }
        case 'rename_sheet': {
          const sheet = resolver.getWorksheet(op.name);
          sheet.name = op.newName;
          resolver.registerRenamed(op.name, op.newName, sheet);
          await context.sync();
          pendingOpsCount = 0;
          break;
        }
        case 'delete_sheet': {
          resolver.getWorksheet(op.name).delete();
          resolver.registerDeleted(op.name);
          await context.sync();
          pendingOpsCount = 0;
          break;
        }
        case 'insert_rows': {
          // Row addresses are 1-based in A1 notation; op.at is 0-based.
          const sheet = resolver.getWorksheet(op.sheet);
          sheet.getRange(`${op.at + 1}:${op.at + op.count}`).insert(Excel.InsertShiftDirection.down);
          await context.sync();
          pendingOpsCount = 0;
          break;
        }
        case 'delete_rows': {
          const sheet = resolver.getWorksheet(op.sheet);
          sheet.getRange(`${op.at + 1}:${op.at + op.count}`).delete(Excel.DeleteShiftDirection.up);
          await context.sync();
          pendingOpsCount = 0;
          break;
        }
        case 'insert_cols': {
          const sheet = resolver.getWorksheet(op.sheet);
          sheet
            .getRange(`${indexToCol(op.at)}:${indexToCol(op.at + op.count - 1)}`)
            .insert(Excel.InsertShiftDirection.right);
          await context.sync();
          pendingOpsCount = 0;
          break;
        }
        case 'delete_cols': {
          const sheet = resolver.getWorksheet(op.sheet);
          sheet
            .getRange(`${indexToCol(op.at)}:${indexToCol(op.at + op.count - 1)}`)
            .delete(Excel.DeleteShiftDirection.left);
          await context.sync();
          pendingOpsCount = 0;
          break;
        }
        case 'sort_range': {
          const range = getRange(op.range, resolver);
          range.sort.apply(
            [{ key: op.keyOffset, ascending: op.ascending }],
            false,
            op.hasHeader,
            Excel.SortOrientation.rows,
          );
          pendingOpsCount++;
          break;
        }
        case 'format_range': {
          const range = getRange(op.range, resolver);
          const f = op.format;
          if (f.numberFormat !== undefined) {
            // numberFormat must be a matrix matching the range dimensions.
            let rowCount: number | null = null;
            let colCount: number | null = null;
            try {
              const parsed = parseA1(op.range, 'Active');
              rowCount = parsed.endRow - parsed.startRow + 1;
              colCount = parsed.endCol - parsed.startCol + 1;
            } catch {
              // Non-standard address syntax, fallback to loading from Excel
            }
            if (rowCount !== null && colCount !== null && rowCount > 0 && colCount > 0) {
              range.numberFormat = Array.from({ length: rowCount }, () =>
                Array.from({ length: colCount! }, () => f.numberFormat!),
              );
            } else {
              // Flush any prior queued ops and read dimensions
              range.load(['rowCount', 'columnCount']);
              await context.sync();
              pendingOpsCount = 0;
              range.numberFormat = Array.from({ length: range.rowCount }, () =>
                Array.from({ length: range.columnCount }, () => f.numberFormat!),
              );
            }
          }
          if (f.bold !== undefined) range.format.font.bold = f.bold;
          if (f.italic !== undefined) range.format.font.italic = f.italic;
          if (f.fontColor !== undefined) range.format.font.color = f.fontColor;
          if (f.fillColor !== undefined) range.format.fill.color = f.fillColor;
          if (f.horizontalAlignment !== undefined) {
            range.format.horizontalAlignment = {
              left: Excel.HorizontalAlignment.left,
              center: Excel.HorizontalAlignment.center,
              right: Excel.HorizontalAlignment.right,
            }[f.horizontalAlignment];
          }
          if (f.fontSize !== undefined) range.format.font.size = f.fontSize;
          if (f.wrapText !== undefined) range.format.wrapText = f.wrapText;
          if (f.border !== undefined) applyBorders(range, f.border);
          if (f.autofitColumns) range.format.autofitColumns();
          pendingOpsCount++;
          break;
        }
        case 'write_range': {
          const range = getRange(op.range, resolver);
          // Setting `formulas` writes formulas for "=..." strings and treats
          // everything else as a literal value — one matrix handles both.
          const matrix = op.cells.map((row) =>
            row.map((cell) => (cell.formula !== undefined ? cell.formula : (cell.value ?? ''))),
          );
          range.formulas = matrix as unknown as string[][];
          pendingOpsCount++;
          break;
        }
      }

      // Flush in batches of BATCH_SIZE to keep IPC messages bounded
      if (pendingOpsCount >= BATCH_SIZE) {
        await context.sync();
        pendingOpsCount = 0;
      }
    }

    // Flush any remaining batched operations at the end
    if (pendingOpsCount > 0) {
      await context.sync();
    }
  });
}

function getEdgeMap(): Record<string, Excel.BorderIndex[]> {
  return {
    top: [Excel.BorderIndex.edgeTop],
    bottom: [Excel.BorderIndex.edgeBottom],
    left: [Excel.BorderIndex.edgeLeft],
    right: [Excel.BorderIndex.edgeRight],
    insideHorizontal: [Excel.BorderIndex.insideHorizontal],
    insideVertical: [Excel.BorderIndex.insideVertical],
    outline: [
      Excel.BorderIndex.edgeTop,
      Excel.BorderIndex.edgeBottom,
      Excel.BorderIndex.edgeLeft,
      Excel.BorderIndex.edgeRight,
    ],
    all: [
      Excel.BorderIndex.edgeTop,
      Excel.BorderIndex.edgeBottom,
      Excel.BorderIndex.edgeLeft,
      Excel.BorderIndex.edgeRight,
      Excel.BorderIndex.insideHorizontal,
      Excel.BorderIndex.insideVertical,
    ],
  };
}

function applyBorders(range: Excel.Range, border: BorderFormat): void {
  const edgeMap = getEdgeMap();
  const style =
    border.style === 'double' ? Excel.BorderLineStyle.double : Excel.BorderLineStyle.continuous;
  const weight =
    border.style === 'thick'
      ? Excel.BorderWeight.thick
      : border.style === 'medium'
        ? Excel.BorderWeight.medium
        : Excel.BorderWeight.thin;
  const indexes = new Set<Excel.BorderIndex>();
  for (const edge of border.edges) for (const idx of edgeMap[edge] ?? []) indexes.add(idx);
  for (const idx of indexes) {
    const b = range.format.borders.getItem(idx);
    b.style = style;
    b.weight = weight;
    if (border.color) b.color = border.color;
  }
}

/** Case-insensitive substring search across sheets, capped at 50 matches. */
export async function findCells(
  query: string,
  sheetName?: string,
): Promise<{ address: string; value: string }[]> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load('items/name');
    await context.sync();
    const cleanTarget = sheetName?.replace(/^'|'$/g, '').trim().toLowerCase();
    const targets = sheets.items.filter(
      (ws) => !cleanTarget || ws.name.trim().toLowerCase() === cleanTarget,
    );
    if (targets.length === 0) throw new Error(`No such sheet: ${sheetName}`);

    const matches: { address: string; value: string }[] = [];
    for (const ws of targets) {
      const found = ws.findAllOrNullObject(query, { completeMatch: false, matchCase: false });
      found.load(['address', 'isNullObject']);
      await context.sync();
      if (found.isNullObject) continue;
      for (const address of found.address.split(',')) {
        // Expand multi-cell areas conservatively: report the area address.
        const range = ws.getRange(address);
        range.load(['values', 'address']);
        await context.sync();
        const flat = range.values.flat();
        matches.push({ address: range.address, value: String(flat[0] ?? '') });
        if (matches.length >= 50) return matches;
      }
    }
    return matches;
  });
}

interface WorksheetResolver {
  getWorksheet(rawName: string): Excel.Worksheet;
  findExisting(rawName: string): Excel.Worksheet | undefined;
  registerAdded(name: string, ws: Excel.Worksheet): void;
  registerRenamed(oldName: string, newName: string, ws: Excel.Worksheet): void;
  registerDeleted(name: string): void;
}

/**
 * Pre-loads the workbook's worksheet collection so range lookups can match
 * sheet names case-insensitively. A name that does not resolve throws, exactly
 * like the emulator: sheet targets are resolved once at staging (and shown in
 * the preview), so execution never guesses — in particular it never falls
 * back to the active worksheet.
 */
async function loadWorksheetResolver(context: Excel.RequestContext): Promise<WorksheetResolver> {
  const sheets = context.workbook.worksheets;
  sheets.load('items/name');
  await context.sync();

  const byExact = new Map<string, Excel.Worksheet>();
  const byLower = new Map<string, Excel.Worksheet>();
  for (const ws of sheets.items ?? []) {
    byExact.set(ws.name, ws);
    byLower.set(ws.name.trim().toLowerCase(), ws);
  }

  const findExisting = (rawName: string): Excel.Worksheet | undefined => {
    const clean = rawName.replace(/^'|'$/g, '').trim();
    if (!clean) return undefined;
    return byExact.get(clean) ?? byLower.get(clean.toLowerCase());
  };

  return {
    findExisting,
    getWorksheet(rawName: string): Excel.Worksheet {
      const ws = findExisting(rawName);
      if (!ws) throw new Error(`No such sheet: ${rawName.replace(/^'|'$/g, '').trim()}`);
      return ws;
    },
    registerAdded(name: string, ws: Excel.Worksheet): void {
      const clean = name.replace(/^'|'$/g, '').trim();
      byExact.set(clean, ws);
      byLower.set(clean.toLowerCase(), ws);
    },
    registerRenamed(oldName: string, newName: string, ws: Excel.Worksheet): void {
      const cleanOld = oldName.replace(/^'|'$/g, '').trim();
      const cleanNew = newName.replace(/^'|'$/g, '').trim();
      byExact.delete(cleanOld);
      byLower.delete(cleanOld.toLowerCase());
      byExact.set(cleanNew, ws);
      byLower.set(cleanNew.toLowerCase(), ws);
    },
    registerDeleted(name: string): void {
      const clean = name.replace(/^'|'$/g, '').trim();
      byExact.delete(clean);
      byLower.delete(clean.toLowerCase());
    },
  };
}

function getRange(ref: string, resolver: WorksheetResolver): Excel.Range {
  const bang = ref.lastIndexOf('!');
  if (bang < 0) throw new Error(`Reference is missing a sheet name: ${ref}`);
  return resolver.getWorksheet(ref.slice(0, bang)).getRange(ref.slice(bang + 1));
}

/**
 * Stable per-workbook identity, stored in the document itself so memory
 * survives renames. Includes a random component generated on first use.
 */
export async function getWorkbookIdentity(): Promise<{ id: string; name: string }> {
  const settings = Office.context.document.settings;
  let id = settings.get('excelai-workbook-id') as string | null;
  if (!id) {
    id = crypto.randomUUID();
    settings.set('excelai-workbook-id', id);
    await new Promise<void>((resolve) => settings.saveAsync(() => resolve()));
  }
  const url = Office.context.document.url ?? 'Workbook';
  const name = url.split(/[\\/]/).pop() || 'Workbook';
  return { id, name };
}

export interface SelectionInfo {
  address: string;
  cellCount: number;
}

/**
 * Reads the user's currently selected range in Excel, if available.
 */
export async function getSelectedRangeInfo(): Promise<SelectionInfo | null> {
  if (typeof Excel === 'undefined') return null;
  try {
    return await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(['address', 'cellCount']);
      await context.sync();
      return { address: range.address, cellCount: range.cellCount };
    });
  } catch {
    return null;
  }
}

/**
 * Registers a listener that triggers whenever the user selects a new range in Excel.
 * Returns an unregister cleanup function.
 */
export function registerSelectionChangeListener(
  onChange: (info: SelectionInfo) => void,
): () => void {
  if (typeof Office === 'undefined' || !Office.context?.document) return () => {};

  let active = true;

  const onSelectionChange = async (): Promise<void> => {
    if (!active) return;
    const info = await getSelectedRangeInfo();
    if (info && active) onChange(info);
  };

  try {
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      onSelectionChange,
    );
  } catch {
    // Non-fatal if host environment doesn't support the event
  }

  return () => {
    active = false;
    try {
      Office.context.document.removeHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        onSelectionChange,
      );
    } catch {
      // Ignore cleanup failure
    }
  };
}

/**
 * Selects/highlights a given range in Excel, jumping the sheet view to it.
 */
export async function selectRangeOnSheet(rangeRef: string): Promise<boolean> {
  if (typeof Excel === 'undefined') return false;
  try {
    return await Excel.run(async (context) => {
      const resolver = await loadWorksheetResolver(context);
      const range = getRange(rangeRef, resolver);
      range.select();
      await context.sync();
      return true;
    });
  } catch {
    return false;
  }
}
