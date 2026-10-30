import type { ChangeOp, ChangePreviewItem, ChangeSetStatus, UndoFidelity } from '@excelai/shared';
import { DiffInspector } from '../../DiffInspector.js';

export interface ChangeSetCardProps {
  id: string;
  preview: ChangePreviewItem[];
  undoFidelity: UndoFidelity;
  status: ChangeSetStatus | 'unknown';
  onApprove: (id: string) => void;
  onReject: (id: string) => void;
  onUndo: (id: string) => void;
}

const TITLES: Record<Exclude<ChangeSetCardProps['status'], 'staged'>, string> = {
  applying: 'Applying…',
  applied: 'Applied',
  rolled_back: 'Not applied (rolled back)',
  failed_partial: 'Partly applied: check the workbook',
  rejected: 'Dismissed',
  undoing: 'Undoing…',
  undone: 'Undone',
  unknown: 'No longer available',
};

/** Ops that make a set impossible to undo, phrased for the card. */
const NOT_UNDOABLE: Partial<Record<ChangeOp['kind'], string>> = {
  insert_rows: 'inserts rows',
  delete_rows: 'deletes rows',
  insert_cols: 'inserts columns',
  delete_cols: 'deletes columns',
  rename_sheet: 'renames a sheet',
  delete_sheet: 'deletes a sheet',
};

export function ChangeSetCard({
  id,
  preview,
  undoFidelity,
  status,
  onApprove,
  onReject,
  onUndo,
}: ChangeSetCardProps): JSX.Element {
  const n = preview.length;
  const sheets = targetSheets(preview);
  const note = status === 'staged' || status === 'applied' ? undoNote(undoFidelity, preview) : null;
  return (
    <div className={`cs ${status}`}>
      <div className="cs-head">
        <span className="cs-title">
          {status === 'staged' ? `${n} proposed change${n === 1 ? '' : 's'}` : TITLES[status]}
        </span>
        <span className="cs-actions">
          {status === 'staged' && (
            <>
              <button className="sm" onClick={() => onApprove(id)}>
                Apply
              </button>
              <button className="sm subtle" onClick={() => onReject(id)}>
                Dismiss
              </button>
            </>
          )}
          {status === 'applied' && undoFidelity !== 'none' && (
            <button className="sm subtle" onClick={() => onUndo(id)}>
              Undo
            </button>
          )}
        </span>
      </div>
      {sheets.length > 0 && (
        <div className="cs-sheet">
          {sheets.length === 1 ? 'Sheet' : 'Sheets'}: {sheets.join(', ')}
        </div>
      )}
      {note && <div className="cs-note">{note}</div>}
      <div className="cs-ops">
        {preview.map((p, i) => (
          <DiffInspector key={i} item={p} />
        ))}
      </div>
    </div>
  );
}

/** What undo will and won't restore, shown before approval and after apply. */
function undoNote(fidelity: UndoFidelity, preview: ChangePreviewItem[]): string | null {
  if (fidelity === 'values_only') return 'Undo restores values and formulas; formatting stays as applied.';
  if (fidelity === 'full') return null;
  const reasons = [...new Set(preview.map(({ op }) => NOT_UNDOABLE[op.kind]).filter(Boolean))];
  return `Can't be undone: it ${reasons.join(', ')}.`;
}

/** Distinct sheets the ops touch, as resolved at staging (exact names). */
function targetSheets(preview: ChangePreviewItem[]): string[] {
  const names = preview.map(({ op }) => {
    switch (op.kind) {
      case 'write_range':
      case 'clear_range':
      case 'format_range':
      case 'sort_range':
        return op.range.slice(0, op.range.lastIndexOf('!')).replace(/^'|'$/g, '');
      case 'add_sheet':
      case 'rename_sheet':
      case 'delete_sheet':
        return op.name;
      default:
        return op.sheet;
    }
  });
  return [...new Set(names.filter((name) => name.length > 0))];
}

export function OpRow({ item }: { item: ChangePreviewItem }): JSX.Element {
  const { op } = item;
  let target: string;
  let desc: string;
  switch (op.kind) {
    case 'write_range': {
      const cells = op.cells.length * (op.cells[0]?.length ?? 0);
      target = op.range;
      desc = cells === 1 ? 'write' : `write ${cells} cells`;
      break;
    }
    case 'clear_range':
      target = op.range;
      desc = 'clear';
      break;
    case 'format_range':
      target = op.range;
      desc = `format: ${Object.keys(op.format).join(', ')}`;
      break;
    case 'sort_range':
      target = op.range;
      desc = `sort ${op.ascending ? '↑' : '↓'}`;
      break;
    case 'add_sheet':
      target = op.name;
      desc = 'new sheet';
      break;
    case 'rename_sheet':
      target = op.name;
      desc = `rename → ${op.newName}`;
      break;
    case 'delete_sheet':
      target = op.name;
      desc = 'delete sheet';
      break;
    case 'insert_rows':
      target = op.sheet;
      desc = `insert ${op.count} row${op.count === 1 ? '' : 's'} at ${op.at + 1}`;
      break;
    case 'delete_rows':
      target = op.sheet;
      desc = `delete ${op.count} row${op.count === 1 ? '' : 's'} at ${op.at + 1}`;
      break;
    case 'insert_cols':
      target = op.sheet;
      desc = `insert ${op.count} col${op.count === 1 ? '' : 's'}`;
      break;
    case 'delete_cols':
      target = op.sheet;
      desc = `delete ${op.count} col${op.count === 1 ? '' : 's'}`;
      break;
  }
  return (
    <li title={`${target} · ${op.reason}`}>
      <code>{target.replace(/^[^!]*!/, '')}</code>
      <span className="cs-desc">{desc}</span>
      <span className="cs-reason">{op.reason}</span>
    </li>
  );
}
