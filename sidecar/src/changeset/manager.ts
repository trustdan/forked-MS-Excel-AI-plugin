import { randomUUID } from 'node:crypto';
import {
  parseA1,
  formatA1,
  rangeCellCount,
  type CellData,
  type ChangeOp,
  type ChangePreviewItem,
  type ChangeSet,
  type RangeSnapshot,
  type SheetTarget,
  type UndoFidelity,
} from '@excelai/shared';
import type { WorkbookExecutor } from '../workbook/executor.js';
import { assertTargetsUnchanged, resolveSheetTargets } from './targets.js';

const MAX_SNAPSHOT_CELLS = 100_000;
const MAX_CELLS_LISTED = 20;

// Range snapshots cannot express these: after rows/cols shift or a sheet is
// renamed or deleted, restoring a snapshot would write to the wrong cells.
const NOT_UNDOABLE: ChangeOp['kind'][] = [
  'rename_sheet',
  'delete_sheet',
  'insert_rows',
  'delete_rows',
  'insert_cols',
  'delete_cols',
];

export interface ApplyResult {
  ok: boolean;
  /**
   * Verification problems found after apply. When there are any, the set is
   * rolled back: its status says whether that fully succeeded.
   */
  problems: string[];
}

/**
 * Stages, previews, applies, verifies and undoes change-sets.
 *
 * Safety contract:
 *  - ops are staged, never applied directly by the agent;
 *  - before apply, affected ranges are snapshotted for undo;
 *  - after apply, touched ranges are re-read and scanned for newly introduced
 *    error values; any regression rolls the change-set back automatically;
 *  - undo refuses to overwrite cells that changed after apply.
 */
export class ChangeSetManager {
  private changeSets = new Map<string, ChangeSet>();
  /** Serializes every workbook mutation made through this manager (one per workbook). */
  private workbookLock: Promise<unknown> = Promise.resolve();
  private inFlight = new Map<string, Promise<unknown>>();

  constructor(
    private executor: WorkbookExecutor,
    private readonly workbookId: string,
  ) {}

  setExecutor(executor: WorkbookExecutor): void {
    this.executor = executor;
  }

  /** Resolves sheet targets against the live workbook; throws on an unknown sheet. */
  async stage(stagedOps: ChangeOp[]): Promise<ChangeSet> {
    const { ops, targets } = resolveSheetTargets(stagedOps, await this.executor.getWorkbookMap());
    const cs: ChangeSet = {
      id: randomUUID(),
      workbookId: this.workbookId,
      createdAt: new Date().toISOString(),
      status: 'staged',
      ops,
      targets,
      undoFidelity: undoFidelityOf(ops),
    };
    this.changeSets.set(cs.id, cs);
    return cs;
  }

  get(id: string): ChangeSet | undefined {
    return this.changeSets.get(id);
  }

  list(): ChangeSet[] {
    return [...this.changeSets.values()];
  }

  async preview(id: string): Promise<ChangePreviewItem[]> {
    const cs = this.mustGet(id, 'staged');
    const items: ChangePreviewItem[] = [];
    for (const op of cs.ops) {
      if (op.kind === 'write_range' || op.kind === 'clear_range' || op.kind === 'sort_range') {
        let before: CellData[][] | undefined;
        try {
          before = await this.executor.readRange(op.range);
        } catch {
          before = undefined; // e.g. range on a sheet this same change-set adds
        }
        items.push({ op, ...(before !== undefined ? { before } : {}) });
      } else {
        items.push({ op });
      }
    }
    return items;
  }

  reject(id: string): void {
    const cs = this.mustGet(id, 'staged');
    cs.status = 'rejected';
  }

  /**
   * Applies a staged set. Status flips to `applying` synchronously, before
   * any await, so a concurrent approve/reject/undo of the same set sees it;
   * a duplicate approval gets the in-flight result. The workbook lock keeps
   * different sets from interleaving snapshot/apply/verify phases.
   */
  apply(id: string): Promise<ApplyResult> {
    const cs = this.changeSets.get(id);
    if (cs?.status === 'applying') {
      const pending = this.inFlight.get(id);
      if (pending) return pending as Promise<ApplyResult>;
    }
    try {
      this.mustGet(id, 'staged').status = 'applying';
    } catch (e) {
      return Promise.reject(e);
    }
    return this.track(id, this.withWorkbookLock(() => this.runApply(cs!)));
  }

  /** Undoes an applied set; same serialization rules as {@link apply}. */
  undo(id: string): Promise<void> {
    const cs = this.changeSets.get(id);
    if (cs?.status === 'undoing') {
      const pending = this.inFlight.get(id);
      if (pending) return pending as Promise<void>;
    }
    try {
      const target = this.mustGet(id, 'applied');
      if (target.undoFidelity === 'none') {
        const blockers = new Set(target.ops.filter((op) => NOT_UNDOABLE.includes(op.kind)).map((op) => op.kind));
        throw new Error(`Change-set ${id} cannot be undone: it contains ${[...blockers].join(', ')}`);
      }
      target.status = 'undoing';
    } catch (e) {
      return Promise.reject(e);
    }
    return this.track(id, this.withWorkbookLock(() => this.runUndo(cs!)));
  }

  private async runApply(cs: ChangeSet): Promise<ApplyResult> {
    // Until ops reach the workbook, a failure leaves the set reviewable.
    let touched = false;
    let sheetsBefore = new Set<string>();
    try {
      // 0. Refuse to run if a target sheet was deleted or renamed since staging.
      const mapBefore = await this.executor.getWorkbookMap();
      assertTargetsUnchanged(cs.targets, mapBefore);
      sheetsBefore = new Set(mapBefore.sheets.map((s) => s.id));

      // 1. Snapshot affected ranges for undo.
      const snapshots: RangeSnapshot[] = [];
      let snapshotCells = 0;
      for (const op of cs.ops) {
        if (op.kind !== 'write_range' && op.kind !== 'clear_range' && op.kind !== 'sort_range') continue;
        const addr = parseA1(op.range);
        snapshotCells += rangeCellCount(addr);
        if (snapshotCells > MAX_SNAPSHOT_CELLS) {
          throw new Error(
            `Change-set touches more than ${MAX_SNAPSHOT_CELLS} cells; refusing to apply without splitting it up`,
          );
        }
        let cells: CellData[][];
        try {
          cells = await this.executor.readRange(op.range);
        } catch {
          continue; // sheet does not exist yet (created earlier in this set)
        }
        snapshots.push({ range: formatA1(addr), cells });
      }
      cs.snapshots = snapshots;

      const beforeErrors = countErrors(snapshots.flatMap((s) => s.cells));

      // 2. Apply.
      touched = true;
      await this.executor.applyOps(cs.ops);

      // 3. Verify: re-read touched ranges and scan for newly introduced errors.
      // The re-read is also what undo later checks for conflicting edits.
      const problems: string[] = [];
      const appliedCells: RangeSnapshot[] = [];
      for (const op of cs.ops) {
        if (op.kind !== 'write_range' && op.kind !== 'clear_range' && op.kind !== 'sort_range') continue;
        let after: CellData[][];
        try {
          after = await this.executor.readRange(op.range);
        } catch {
          continue;
        }
        appliedCells.push({ range: formatA1(parseA1(op.range)), cells: after });
        for (const [r, row] of after.entries()) {
          for (const [c, cell] of row.entries()) {
            if (typeof cell.value === 'string' && cell.value.startsWith('#')) {
              // Writing an error value intentionally is almost never what the
              // user asked for; surface every error in the written range.
              problems.push(`${op.range} row ${r + 1} col ${c + 1} now contains ${cell.value}`);
            }
          }
        }
      }

      if (problems.length > 0 && beforeErrors === 0) {
        // Regression introduced by this change-set: roll back.
        await this.rollback(cs, sheetsBefore);
        return { ok: false, problems };
      }

      const added = await this.findAddedSheets(cs, sheetsBefore);
      if (added === undefined) throw new Error('Could not read the workbook after applying the change-set');
      cs.appliedCells = appliedCells;
      cs.addedSheets = added;
      cs.status = 'applied';
      return { ok: true, problems: [] };
    } catch (e) {
      if (touched) await this.rollback(cs, sheetsBefore);
      else cs.status = 'staged';
      throw e;
    }
  }

  /**
   * Refuses (leaving the set `applied`) if anything the set wrote or added
   * changed since apply; otherwise reverts it and confirms by re-reading.
   */
  private async runUndo(cs: ChangeSet): Promise<void> {
    let conflicts: string[];
    try {
      conflicts = await this.findConflicts(cs);
    } catch (e) {
      cs.status = 'applied';
      throw e;
    }
    if (conflicts.length > 0) {
      cs.status = 'applied';
      throw new Error(
        `Undo refused: ${listCells(conflicts)} changed after the change-set was applied, ` +
          'and undo would overwrite them. Nothing was changed.',
      );
    }
    const unrecovered = await this.revert(cs, 'undo', cs.addedSheets ?? []);
    if (unrecovered.length > 0) {
      cs.unrecovered = unrecovered;
      cs.status = 'failed_partial';
      throw new Error('Undo did not complete.');
    }
    cs.status = 'undone';
  }

  /** Cells and added sheets whose content differs from what apply left. */
  private async findConflicts(cs: ChangeSet): Promise<string[]> {
    const conflicts: string[] = [];
    for (const applied of cs.appliedCells ?? []) {
      let now: CellData[][] | undefined;
      try {
        now = await this.executor.readRange(applied.range);
      } catch {
        now = undefined;
      }
      conflicts.push(...mismatchedCells(applied, now));
    }
    if (!cs.addedSheets?.length) return conflicts;

    // Deleting an added sheet must not take anything but this set's own cells with it.
    const map = await this.executor.getWorkbookMap();
    for (const sheet of cs.addedSheets) {
      const now = map.sheets.find((s) => s.id === sheet.id);
      if (!now) {
        conflicts.push(`sheet "${sheet.name}" (deleted)`);
        continue;
      }
      if (!now.usedRange) continue;
      const used = parseA1(now.usedRange);
      if (rangeCellCount(used) > MAX_SNAPSHOT_CELLS) {
        conflicts.push(`sheet "${now.name}" (too much content to check)`);
        continue;
      }
      const ours = (cs.appliedCells ?? []).map((a) => parseA1(a.range)).filter((a) => a.sheet === now.name);
      const cells = await this.executor.readRange(now.usedRange);
      for (const [r, row] of cells.entries()) {
        for (const [c, cell] of row.entries()) {
          const row0 = used.startRow + r;
          const col0 = used.startCol + c;
          const empty = cell.formula === undefined && (cell.value === null || cell.value === '');
          const inOurs = ours.some(
            (a) => row0 >= a.startRow && row0 <= a.endRow && col0 >= a.startCol && col0 <= a.endCol,
          );
          if (!empty && !inOurs) conflicts.push(cellA1(now.name, row0, col0));
        }
      }
    }
    return conflicts;
  }

  /**
   * Reverts the set and ends `rolled_back` only when every cell matches its
   * snapshot, every added sheet is gone and nothing unsnapshotted (formatting,
   * structure) may have changed; otherwise `failed_partial`, listing what
   * could not be confirmed. Never throws.
   */
  private async rollback(cs: ChangeSet, sheetsBefore: Set<string>): Promise<void> {
    const added = await this.findAddedSheets(cs, sheetsBefore);
    const unrecovered = await this.revert(cs, 'rollback', added ?? []);
    const addedNames = addedSheetNames(cs.ops);
    for (const op of cs.ops) {
      const leftBehind =
        op.kind === 'add_sheet'
          ? added === undefined
          : NOT_UNDOABLE.includes(op.kind) || changesExistingFormatting(op, addedNames);
      if (leftBehind) unrecovered.push(describeStructural(op));
    }
    if (unrecovered.length > 0) {
      cs.unrecovered = unrecovered;
      cs.status = 'failed_partial';
    } else {
      cs.status = 'rolled_back';
    }
  }

  /**
   * Restores every snapshot and deletes the given added sheets, then re-reads
   * to confirm. Returns what could not be confirmed restored. Never throws.
   */
  private async revert(
    cs: ChangeSet,
    purpose: 'rollback' | 'undo',
    addedSheets: SheetTarget[],
  ): Promise<string[]> {
    const unrecovered: string[] = [];
    for (const snap of cs.snapshots ?? []) {
      try {
        await this.executor.applyOps([restoreOp(snap, `${purpose}: restore snapshot`)]);
      } catch {
        // The re-read below reports whatever stayed changed.
      }
      let now: CellData[][] | undefined;
      try {
        now = await this.executor.readRange(snap.range);
      } catch {
        now = undefined;
      }
      unrecovered.push(...mismatchedCells(snap, now));
    }
    if (addedSheets.length === 0) return unrecovered;

    const liveSheets = async () => (await this.executor.getWorkbookMap().catch(() => undefined))?.sheets;
    const before = await liveSheets();
    for (const sheet of addedSheets) {
      const now = before?.find((s) => s.id === sheet.id);
      if (!now) continue;
      try {
        await this.executor.applyOps([
          { kind: 'delete_sheet', name: now.name, reason: `${purpose}: remove added sheet` },
        ]);
      } catch {
        // Confirmed below.
      }
    }
    const after = await liveSheets();
    for (const sheet of addedSheets) {
      if (after === undefined || after.some((s) => s.id === sheet.id)) {
        unrecovered.push(`add_sheet "${sheet.name}" (not removed)`);
      }
    }
    return unrecovered;
  }

  /** Sheets that exist now but did not before apply; undefined if the workbook can't be read. */
  private async findAddedSheets(cs: ChangeSet, sheetsBefore: Set<string>): Promise<SheetTarget[] | undefined> {
    if (!cs.ops.some((op) => op.kind === 'add_sheet')) return [];
    try {
      const map = await this.executor.getWorkbookMap();
      return map.sheets.filter((s) => !sheetsBefore.has(s.id)).map(({ id, name }) => ({ id, name }));
    } catch {
      return undefined;
    }
  }

  private withWorkbookLock<T>(fn: () => Promise<T>): Promise<T> {
    const run = this.workbookLock.then(fn);
    this.workbookLock = run.catch(() => undefined);
    return run;
  }

  private track<T>(id: string, run: Promise<T>): Promise<T> {
    this.inFlight.set(id, run);
    void run.then(
      () => this.inFlight.delete(id),
      () => this.inFlight.delete(id),
    );
    return run;
  }

  private mustGet(id: string, expected: ChangeSet['status']): ChangeSet {
    const cs = this.changeSets.get(id);
    if (!cs) throw new Error(`Unknown change-set: ${id}`);
    if (cs.status === 'applying' || cs.status === 'undoing') {
      throw new Error(`Change-set ${id} is already in progress (${cs.status})`);
    }
    if (cs.status !== expected) {
      throw new Error(`Change-set ${id} is ${cs.status}, expected ${expected}`);
    }
    return cs;
  }
}

function undoFidelityOf(ops: ChangeOp[]): UndoFidelity {
  if (ops.some((op) => NOT_UNDOABLE.includes(op.kind))) return 'none';
  const addedNames = addedSheetNames(ops);
  return ops.some((op) => changesExistingFormatting(op, addedNames)) ? 'values_only' : 'full';
}

function addedSheetNames(ops: ChangeOp[]): Set<string> {
  return new Set(ops.flatMap((op) => (op.kind === 'add_sheet' ? [op.name.toLowerCase()] : [])));
}

/**
 * Snapshots hold values and formulas only, so formatting a range (or sorting
 * it, which moves formatting) can't be undone, unless the set adds the sheet:
 * then the formatting goes with it.
 */
function changesExistingFormatting(op: ChangeOp, addedNames: Set<string>): boolean {
  return (
    (op.kind === 'format_range' || op.kind === 'sort_range') &&
    !addedNames.has(parseA1(op.range).sheet.toLowerCase())
  );
}

function restoreOp(snap: RangeSnapshot, reason: string): ChangeOp {
  return {
    kind: 'write_range',
    range: snap.range,
    cells: snap.cells.map((row) =>
      row.map((cell) => (cell.formula !== undefined ? { formula: cell.formula } : { value: cell.value })),
    ),
    reason,
  };
}

function cellA1(sheet: string, row: number, col: number): string {
  return formatA1({ sheet, startRow: row, startCol: col, endRow: row, endCol: col });
}

function listCells(cells: string[]): string {
  const more = cells.length - MAX_CELLS_LISTED;
  return cells.slice(0, MAX_CELLS_LISTED).join(', ') + (more > 0 ? ` and ${more} more` : '');
}

/** Addresses of cells that do not hold `expected`'s content (all of them if unreadable). */
function mismatchedCells(expected: RangeSnapshot, now: CellData[][] | undefined): string[] {
  const addr = parseA1(expected.range);
  const out: string[] = [];
  for (const [r, row] of expected.cells.entries()) {
    for (const [c, want] of row.entries()) {
      const got = now?.[r]?.[c];
      const matches =
        got !== undefined &&
        (want.formula !== undefined
          ? got.formula === want.formula
          : got.formula === undefined && got.value === want.value);
      if (!matches) out.push(cellA1(addr.sheet, addr.startRow + r, addr.startCol + c));
    }
  }
  return out;
}

function describeStructural(op: ChangeOp): string {
  switch (op.kind) {
    case 'add_sheet':
    case 'delete_sheet':
      return `${op.kind} "${op.name}" (not reverted)`;
    case 'rename_sheet':
      return `rename_sheet "${op.name}" → "${op.newName}" (not reverted)`;
    case 'format_range':
      return `format_range ${op.range} (not reverted)`;
    case 'sort_range':
      return `sort_range ${op.range} (formatting not reverted)`;
    case 'insert_rows':
    case 'delete_rows':
    case 'insert_cols':
    case 'delete_cols':
      return `${op.kind} ${op.sheet} at ${op.at} ×${op.count} (not reverted)`;
    default:
      return `${op.kind} (not reverted)`;
  }
}

function countErrors(cells: CellData[][]): number {
  let n = 0;
  for (const row of cells) {
    for (const cell of row) {
      if (typeof cell.value === 'string' && cell.value.startsWith('#')) n++;
    }
  }
  return n;
}
