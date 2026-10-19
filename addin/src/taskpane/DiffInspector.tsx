import React, { useState } from 'react';
import {
  indexToCol,
  parseA1,
  type ChangeOp,
  type ChangePreviewItem,
  type CellData,
} from '@excelai/shared';
import { selectRangeOnSheet } from '../officeBridge.js';

interface DiffInspectorProps {
  item: ChangePreviewItem;
}

interface CellDiff {
  address: string;
  beforeValue?: string;
  beforeFormula?: string;
  afterValue?: string;
  afterFormula?: string;
  isChanged: boolean;
}

/**
 * Visual Before / After Diff Inspector (UI-02)
 * Renders cell-level before/after matrices, format breakdowns, and sheet target selectors.
 */
export const DiffInspector: React.FC<DiffInspectorProps> = ({ item }) => {
  const [expanded, setExpanded] = useState(false);
  const { op, before } = item;

  const target = getTargetString(op);
  const desc = getOpDescription(op);

  const jumpToRange = (e: React.MouseEvent): void => {
    e.stopPropagation();
    if (op.kind === 'write_range' || op.kind === 'clear_range' || op.kind === 'format_range' || op.kind === 'sort_range') {
      void selectRangeOnSheet(op.range);
    }
  };

  return (
    <div className={`diff-inspector ${expanded ? 'expanded' : ''}`}>
      <div className="diff-header" onClick={() => setExpanded(!expanded)} role="button" tabIndex={0}>
        <div className="diff-title-row">
          <span className="diff-chevron">{expanded ? '▼' : '▶'}</span>
          <code
            className="diff-target-badge"
            title="Click to select this range on the sheet"
            onClick={jumpToRange}
          >
            {target.replace(/^[^!]*!/, '')}
          </code>
          <span className="diff-desc">{desc}</span>
        </div>
        <div className="diff-reason">{op.reason}</div>
      </div>

      {expanded && (
        <div className="diff-details">
          {op.kind === 'write_range' && renderWriteDiff(op, before)}
          {op.kind === 'clear_range' && renderClearDiff(op, before)}
          {op.kind === 'format_range' && renderFormatDiff(op)}
          {op.kind === 'sort_range' && renderSortDiff(op)}
          {op.kind !== 'write_range' &&
            op.kind !== 'clear_range' &&
            op.kind !== 'format_range' &&
            op.kind !== 'sort_range' &&
            renderStructuralDiff(op)}
        </div>
      )}
    </div>
  );
};

/* ---------------- Helper: Target & Description ---------------- */

function getTargetString(op: ChangeOp): string {
  switch (op.kind) {
    case 'write_range':
    case 'clear_range':
    case 'format_range':
    case 'sort_range':
      return op.range;
    case 'add_sheet':
    case 'rename_sheet':
    case 'delete_sheet':
      return op.name;
    case 'insert_rows':
    case 'delete_rows':
    case 'insert_cols':
    case 'delete_cols':
      return op.sheet;
  }
}

function getOpDescription(op: ChangeOp): string {
  switch (op.kind) {
    case 'write_range': {
      const cells = op.cells.length * (op.cells[0]?.length ?? 0);
      return cells === 1 ? 'write 1 cell' : `write ${cells} cells`;
    }
    case 'clear_range':
      return 'clear contents';
    case 'format_range':
      return `format (${Object.keys(op.format).join(', ')})`;
    case 'sort_range':
      return `sort ${op.ascending ? 'ascending' : 'descending'}`;
    case 'add_sheet':
      return 'new sheet';
    case 'rename_sheet':
      return `rename → ${op.newName}`;
    case 'delete_sheet':
      return 'delete sheet';
    case 'insert_rows':
      return `insert ${op.count} row${op.count === 1 ? '' : 's'} at row ${op.at + 1}`;
    case 'delete_rows':
      return `delete ${op.count} row${op.count === 1 ? '' : 's'} at row ${op.at + 1}`;
    case 'insert_cols':
      return `insert ${op.count} column${op.count === 1 ? '' : 's'}`;
    case 'delete_cols':
      return `delete ${op.count} column${op.count === 1 ? '' : 's'}`;
  }
}

/* ---------------- Write Diff Renderer ---------------- */

function renderWriteDiff(
  op: Extract<ChangeOp, { kind: 'write_range' }>,
  before?: CellData[][],
): JSX.Element {
  let diffs: CellDiff[] = [];

  try {
    const addr = parseA1(op.range);
    const rowCount = op.cells.length;
    const colCount = op.cells[0]?.length ?? 0;

    for (let r = 0; r < rowCount; r++) {
      for (let c = 0; c < colCount; c++) {
        const cellAddress = `${indexToCol(addr.startCol + c)}${addr.startRow + r + 1}`;
        const bCell = before?.[r]?.[c];
        const aCell = op.cells[r]?.[c];

        const bFormula = bCell?.formula;
        const bVal = bCell?.value !== undefined && bCell.value !== null ? String(bCell.value) : '';
        const aFormula = aCell?.formula;
        const aVal = aCell?.value !== undefined && aCell.value !== null ? String(aCell.value) : '';

        const beforeDisplay = bFormula || bVal;
        const afterDisplay = aFormula || aVal;
        const isChanged = beforeDisplay !== afterDisplay;

        diffs.push({
          address: cellAddress,
          beforeValue: bVal || undefined,
          beforeFormula: bFormula,
          afterValue: aVal || undefined,
          afterFormula: aFormula,
          isChanged,
        });
      }
    }
  } catch {
    // If range parsing fails, fallback
    diffs = [];
  }

  // Filter to changed cells, or show up to 20 cells if all match
  const changedDiffs = diffs.filter((d) => d.isChanged);
  const displayDiffs = changedDiffs.length > 0 ? changedDiffs : diffs.slice(0, 15);
  const maxDisplay = 25;
  const truncated = displayDiffs.length > maxDisplay;
  const visible = displayDiffs.slice(0, maxDisplay);

  return (
    <div className="diff-matrix-wrap">
      <div className="diff-matrix-header">
        <span>Cell-level Changes</span>
        {changedDiffs.length > 0 ? (
          <span className="diff-count-badge">
            {changedDiffs.length} cell{changedDiffs.length === 1 ? '' : 's'} modified
          </span>
        ) : (
          <span className="diff-count-badge muted">No value change</span>
        )}
      </div>
      <table className="diff-table">
        <thead>
          <tr>
            <th>Cell</th>
            <th>Original (Before)</th>
            <th>Proposed (After)</th>
          </tr>
        </thead>
        <tbody>
          {visible.map((d, i) => (
            <tr key={i} className={d.isChanged ? 'row-changed' : 'row-unchanged'}>
              <td className="cell-coord"><code>{d.address}</code></td>
              <td className="cell-before">
                {d.beforeFormula ? (
                  <span className="cell-fx" title={d.beforeFormula}>fx {d.beforeFormula}</span>
                ) : d.beforeValue ? (
                  <span className="cell-val">{d.beforeValue}</span>
                ) : (
                  <span className="cell-empty">(empty)</span>
                )}
              </td>
              <td className="cell-after">
                {d.afterFormula ? (
                  <span className="cell-fx new" title={d.afterFormula}>fx {d.afterFormula}</span>
                ) : d.afterValue ? (
                  <span className="cell-val new">{d.afterValue}</span>
                ) : (
                  <span className="cell-empty">(empty)</span>
                )}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      {truncated && (
        <div className="diff-truncated-hint">
          + {displayDiffs.length - maxDisplay} more cell changes…
        </div>
      )}
    </div>
  );
}

/* ---------------- Clear Diff Renderer ---------------- */

function renderClearDiff(
  _op: Extract<ChangeOp, { kind: 'clear_range' }>,
  before?: CellData[][],
): JSX.Element {
  const nonEmpty = (before ?? []).flat().filter((c) => c.value !== null && c.value !== '');
  return (
    <div className="diff-clear-info">
      <span className="diff-clear-pill">Clear all values & formulas</span>
      {nonEmpty.length > 0 && (
        <span className="diff-clear-note">({nonEmpty.length} currently populated cells will be wiped)</span>
      )}
    </div>
  );
}

/* ---------------- Format Diff Renderer ---------------- */

function renderFormatDiff(op: Extract<ChangeOp, { kind: 'format_range' }>): JSX.Element {
  const f = op.format;
  const badges: { label: string; value: string; color?: string }[] = [];

  if (f.numberFormat) badges.push({ label: 'Format', value: f.numberFormat });
  if (f.bold !== undefined) badges.push({ label: 'Bold', value: f.bold ? 'Yes' : 'No' });
  if (f.italic !== undefined) badges.push({ label: 'Italic', value: f.italic ? 'Yes' : 'No' });
  if (f.fontSize) badges.push({ label: 'Font Size', value: `${f.fontSize}pt` });
  if (f.fontColor) badges.push({ label: 'Font Color', value: f.fontColor, color: f.fontColor });
  if (f.fillColor) badges.push({ label: 'Fill Color', value: f.fillColor, color: f.fillColor });
  if (f.horizontalAlignment) badges.push({ label: 'Align', value: f.horizontalAlignment });
  if (f.wrapText !== undefined) badges.push({ label: 'Wrap Text', value: f.wrapText ? 'On' : 'Off' });
  if (f.autofitColumns) badges.push({ label: 'Autofit', value: 'Enabled' });
  if (f.border) {
    badges.push({
      label: 'Borders',
      value: `${f.border.edges.join(', ')} (${f.border.style ?? 'thin'})`,
    });
  }

  return (
    <div className="diff-format-grid">
      {badges.map((b, i) => (
        <div key={i} className="diff-format-item">
          <span className="diff-format-label">{b.label}:</span>
          <span className="diff-format-value">
            {b.color && (
              <span className="diff-color-swatch" style={{ backgroundColor: b.color }} />
            )}
            {b.value}
          </span>
        </div>
      ))}
    </div>
  );
}

/* ---------------- Sort Diff Renderer ---------------- */

function renderSortDiff(op: Extract<ChangeOp, { kind: 'sort_range' }>): JSX.Element {
  return (
    <div className="diff-sort-info">
      <span>Sort order: <strong>{op.ascending ? 'Ascending (A → Z / 1 → 9)' : 'Descending (Z → A / 9 → 1)'}</strong></span>
      {op.hasHeader && <span className="diff-badge">First row preserved as header</span>}
    </div>
  );
}

/* ---------------- Structural Diff Renderer ---------------- */

function renderStructuralDiff(op: ChangeOp): JSX.Element {
  return (
    <div className="diff-structural-info">
      <span className="diff-badge-subtle">{getOpDescription(op)}</span>
    </div>
  );
}
