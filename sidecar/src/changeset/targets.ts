import {
  formatA1,
  parseA1,
  type ChangeOp,
  type SheetTarget,
  type WorkbookMap,
} from '@excelai/shared';

interface LiveSheet {
  /** Name at this point in the op sequence. */
  name: string;
  /** null for sheets this change-set creates. */
  id: string | null;
  /** Name in the workbook before any op ran. */
  originalName: string;
}

const cleanName = (raw: string): string => raw.replace(/^'|'$/g, '').trim();
const sheetKey = (raw: string): string => cleanName(raw).toLowerCase();

/**
 * Resolves every sheet an op sequence refers to against the workbook, once,
 * at staging time. Names match case-insensitively and are rewritten to the
 * exact worksheet name, so the preview shows the real target. The set's own
 * add/rename/delete ops are replayed in order, so later ops may target sheets
 * it creates. Throws on any sheet that does not exist — never guesses.
 */
export function resolveSheetTargets(
  ops: ChangeOp[],
  map: WorkbookMap,
): { ops: ChangeOp[]; targets: SheetTarget[] } {
  const live = new Map<string, LiveSheet>();
  for (const s of map.sheets) live.set(sheetKey(s.name), { name: s.name, id: s.id, originalName: s.name });
  const targets = new Map<string, SheetTarget>();

  const resolve = (raw: string): LiveSheet => {
    const sheet = live.get(sheetKey(raw));
    if (!sheet) {
      const names = [...live.values()].map((s) => s.name).join(', ');
      throw new Error(`No such sheet: ${cleanName(raw)}. Sheets in this workbook: ${names || '(none)'}`);
    }
    if (sheet.id !== null) targets.set(sheet.id, { id: sheet.id, name: sheet.originalName });
    return sheet;
  };
  const resolveRange = (range: string): string => {
    const addr = parseA1(range);
    return formatA1({ ...addr, sheet: resolve(addr.sheet).name });
  };
  const assertFree = (raw: string): void => {
    const existing = live.get(sheetKey(raw));
    if (existing) throw new Error(`Sheet already exists: ${existing.name}`);
  };

  const resolved = ops.map((op): ChangeOp => {
    switch (op.kind) {
      case 'write_range':
      case 'clear_range':
      case 'format_range':
      case 'sort_range':
        return { ...op, range: resolveRange(op.range) };
      case 'insert_rows':
      case 'delete_rows':
      case 'insert_cols':
      case 'delete_cols':
        return { ...op, sheet: resolve(op.sheet).name };
      case 'add_sheet': {
        assertFree(op.name);
        const name = cleanName(op.name);
        live.set(sheetKey(name), { name, id: null, originalName: name });
        return { ...op, name };
      }
      case 'rename_sheet': {
        const sheet = resolve(op.name);
        const oldName = sheet.name;
        const newName = cleanName(op.newName);
        if (sheetKey(newName) !== sheetKey(oldName)) assertFree(newName);
        live.delete(sheetKey(oldName));
        live.set(sheetKey(newName), { ...sheet, name: newName });
        return { ...op, name: oldName, newName };
      }
      case 'delete_sheet': {
        const sheet = resolve(op.name);
        live.delete(sheetKey(sheet.name));
        return { ...op, name: sheet.name };
      }
    }
  });

  return { ops: resolved, targets: [...targets.values()] };
}

/**
 * Fails explicitly if a target sheet was deleted or renamed after staging,
 * instead of letting the ops land on whatever now carries that name.
 */
export function assertTargetsUnchanged(targets: SheetTarget[], map: WorkbookMap): void {
  for (const t of targets) {
    const now = map.sheets.find((s) => s.id === t.id);
    if (!now) {
      throw new Error(`Sheet "${t.name}" no longer exists. The change-set was not applied.`);
    }
    if (now.name !== t.name) {
      throw new Error(
        `Sheet "${t.name}" was renamed to "${now.name}" after the change-set was proposed. The change-set was not applied; ask again to target the renamed sheet.`,
      );
    }
  }
}
