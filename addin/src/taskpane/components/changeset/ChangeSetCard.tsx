import type { ChangePreviewItem } from '@excelai/shared';
import { DiffInspector } from '../../DiffInspector.js';

export interface ChangeSetCardProps {
  id: string;
  preview: ChangePreviewItem[];
  status: 'staged' | 'applied' | 'rejected' | 'undone';
  onApprove: (id: string) => void;
  onReject: (id: string) => void;
  onUndo: (id: string) => void;
}

export function ChangeSetCard({
  id,
  preview,
  status,
  onApprove,
  onReject,
  onUndo,
}: ChangeSetCardProps): JSX.Element {
  const n = preview.length;
  return (
    <div className={`cs ${status}`}>
      <div className="cs-head">
        <span className="cs-title">
          {status === 'staged'
            ? `${n} proposed change${n === 1 ? '' : 's'}`
            : status === 'applied'
              ? 'Applied'
              : status === 'undone'
                ? 'Undone'
                : 'Dismissed'}
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
          {status === 'applied' && (
            <button className="sm subtle" onClick={() => onUndo(id)}>
              Undo
            </button>
          )}
        </span>
      </div>
      <div className="cs-ops">
        {preview.map((p, i) => (
          <DiffInspector key={i} item={p} />
        ))}
      </div>
    </div>
  );
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
