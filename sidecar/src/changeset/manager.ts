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
} from '@excelai/shared';
import type { WorkbookExecutor } from '../workbook/executor.js';
import { assertTargetsUnchanged, resolveSheetTargets } from './targets.js';

const MAX_SNAPSHOT_CELLS = 100_000;

export interface ApplyResult {
  ok: boolean;
  /** Verification problems found after apply (apply is rolled back if any). */
  problems: string[];
}

/**
 * Stages, previews, applies, verifies and undoes change-sets.
 *
 * Safety contract:
 *  - ops are staged, never applied directly by the agent;
 *  - before apply, affected ranges are snapshotted for undo;
 *  - after apply, touched ranges are re-read and scanned for newly introduced
 *    error values; any regression rolls the change-set back automatically.
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
    // Structural/format ops are not snapshot-restorable in v1: undo restores
    // values and formulas exactly, but not sheet structure or formatting.
    const NON_RESTORABLE: ChangeOp['kind'][] = [
      'add_sheet',
      'rename_sheet',
      'delete_sheet',
      'format_range',
      'insert_rows',
      'delete_rows',
      'insert_cols',
      'delete_cols',
    ];
    const hasStructural = ops.some((op) => NON_RESTORABLE.includes(op.kind));
    const cs: ChangeSet = {
      id: randomUUID(),
      workbookId: this.workbookId,
      createdAt: new Date().toISOString(),
      status: 'staged',
      ops,
      targets,
      undoFidelity: hasStructural ? 'partial' : 'full',
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
      this.mustGet(id, 'applied').status = 'undoing';
    } catch (e) {
      return Promise.reject(e);
    }
    return this.track(
      id,
      this.withWorkbookLock(async () => {
        try {
          await this.restoreSnapshots(cs!);
        } catch (e) {
          cs!.status = 'failed_partial';
          throw e;
        }
        cs!.status = 'undone';
      }),
    );
  }

  private async runApply(cs: ChangeSet): Promise<ApplyResult> {
    // Until ops reach the workbook, a failure leaves the set reviewable.
    let touched = false;
    try {
      // 0. Refuse to run if a target sheet was deleted or renamed since staging.
      if (cs.targets.length > 0) {
        assertTargetsUnchanged(cs.targets, await this.executor.getWorkbookMap());
      }

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
      const problems: string[] = [];
      for (const op of cs.ops) {
        if (op.kind !== 'write_range' && op.kind !== 'clear_range' && op.kind !== 'sort_range') continue;
        let after: CellData[][];
        try {
          after = await this.executor.readRange(op.range);
        } catch {
          continue;
        }
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
        await this.restoreSnapshots(cs);
        cs.status = 'staged';
        return { ok: false, problems };
      }

      cs.status = 'applied';
      return { ok: true, problems: [] };
    } catch (e) {
      // E2b adds rollback here; until then, say the workbook may be half-changed.
      cs.status = touched ? 'failed_partial' : 'staged';
      throw e;
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

  private async restoreSnapshots(cs: ChangeSet): Promise<void> {
    const ops: ChangeOp[] = [];
    for (const snap of cs.snapshots ?? []) {
      ops.push({
        kind: 'write_range',
        range: snap.range,
        cells: snap.cells.map((row) =>
          row.map((cell) =>
            cell.formula !== undefined ? { formula: cell.formula } : { value: cell.value },
          ),
        ),
        reason: 'undo: restore snapshot',
      });
    }
    if (ops.length > 0) await this.executor.applyOps(ops);
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

function countErrors(cells: CellData[][]): number {
  let n = 0;
  for (const row of cells) {
    for (const cell of row) {
      if (typeof cell.value === 'string' && cell.value.startsWith('#')) n++;
    }
  }
  return n;
}
