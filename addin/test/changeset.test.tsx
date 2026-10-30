import { describe, expect, it, vi } from 'vitest';
import { render, screen, fireEvent } from '@testing-library/react';
import { ChangeSetCard } from '../src/taskpane/components/changeset/ChangeSetCard.js';

describe('ChangeSetCard', () => {
  const mockPreview = [
    {
      op: {
        kind: 'write_range' as const,
        range: 'Sheet1!A1',
        cells: [[{ formula: '=SUM(B1:B10)' }]],
        reason: 'Calculate column total',
      },
      before: [[{ value: 0 }]],
      after: [[{ value: null, formula: '=SUM(B1:B10)' }]],
    },
  ];

  it('renders staged change-set with Apply and Dismiss buttons', () => {
    const onApprove = vi.fn();
    const onReject = vi.fn();
    const onUndo = vi.fn();

    render(
      <ChangeSetCard
        id="cs-1"
        preview={mockPreview}
        undoFidelity="full"
        status="staged"
        onApprove={onApprove}
        onReject={onReject}
        onUndo={onUndo}
      />,
    );

    expect(screen.getByText('1 proposed change')).toBeDefined();

    const applyBtn = screen.getByRole('button', { name: 'Apply' });
    const dismissBtn = screen.getByRole('button', { name: 'Dismiss' });

    expect(applyBtn).toBeDefined();
    expect(dismissBtn).toBeDefined();

    fireEvent.click(applyBtn);
    expect(onApprove).toHaveBeenCalledWith('cs-1');

    fireEvent.click(dismissBtn);
    expect(onReject).toHaveBeenCalledWith('cs-1');
  });

  it('renders applied change-set with Undo button', () => {
    const onUndo = vi.fn();

    render(
      <ChangeSetCard
        id="cs-2"
        preview={mockPreview}
        undoFidelity="full"
        status="applied"
        onApprove={() => {}}
        onReject={() => {}}
        onUndo={onUndo}
      />,
    );

    expect(screen.getByText('Applied')).toBeDefined();
    const undoBtn = screen.getByRole('button', { name: 'Undo' });
    expect(undoBtn).toBeDefined();

    fireEvent.click(undoBtn);
    expect(onUndo).toHaveBeenCalledWith('cs-2');
  });

  it('shows the sheet each staged change targets', () => {
    render(
      <ChangeSetCard
        id="cs-3"
        preview={[
          { op: { kind: 'clear_range', range: "'2025 Payroll'!A1:B2", reason: 'reset' } },
          { op: { kind: 'add_sheet', name: 'Summary', reason: 'new output' } },
        ]}
        undoFidelity="full"
        status="staged"
        onApprove={() => {}}
        onReject={() => {}}
        onUndo={() => {}}
      />,
    );
    expect(screen.getByText('Sheets: 2025 Payroll, Summary')).toBeDefined();
  });

  const noop = () => {};

  it('offers no Undo for a set that cannot be undone, and says why before approval', () => {
    const preview = [
      { op: { kind: 'insert_rows' as const, sheet: 'Sheet1', at: 0, count: 1, reason: 'room' } },
      { op: { kind: 'rename_sheet' as const, name: 'Sheet1', newName: 'Data', reason: 'clarity' } },
    ];
    const props = { id: 'cs-4', preview, undoFidelity: 'none' as const, onApprove: noop, onReject: noop, onUndo: noop };

    const { rerender } = render(<ChangeSetCard {...props} status="staged" />);
    expect(screen.getByText("Can't be undone: it inserts rows, renames a sheet.")).toBeDefined();

    rerender(<ChangeSetCard {...props} status="applied" />);
    expect(screen.queryByRole('button', { name: 'Undo' })).toBeNull();
  });

  it('says formatting stays when undo restores values only', () => {
    render(
      <ChangeSetCard
        id="cs-5"
        preview={mockPreview}
        undoFidelity="values_only"
        status="applied"
        onApprove={noop}
        onReject={noop}
        onUndo={noop}
      />,
    );
    expect(screen.getByRole('button', { name: 'Undo' })).toBeDefined();
    expect(screen.getByText('Undo restores values and formulas; formatting stays as applied.')).toBeDefined();
  });

  it.each([
    ['rolled_back', 'Not applied (rolled back)'],
    ['failed_partial', 'Partly applied: check the workbook'],
    ['unknown', 'No longer available'],
  ] as const)('shows %s with no actions', (status, title) => {
    render(
      <ChangeSetCard
        id="cs-6"
        preview={mockPreview}
        undoFidelity="full"
        status={status}
        onApprove={noop}
        onReject={noop}
        onUndo={noop}
      />,
    );
    expect(screen.getByText(title)).toBeDefined();
    for (const name of ['Apply', 'Dismiss', 'Undo']) {
      expect(screen.queryByRole('button', { name })).toBeNull();
    }
  });
});

describe('officeBridge sheet targeting', () => {
  /** Minimal Excel.run fake over the given worksheets; the first one is active. */
  function installExcel(sheetNames: string[]) {
    const written = new Map<string, unknown[][]>();
    const worksheets = sheetNames.map((name) => ({
      name,
      getRange: vi.fn((address: string) => {
        const range = {
          values: [[1]],
          formulas: [[1]],
          load: vi.fn(),
        };
        Object.defineProperty(range, 'formulas', {
          get: () => [[1]],
          set: (v: unknown[][]) => written.set(`${name}!${address}`, v),
        });
        return range;
      }),
    }));
    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (cb: (ctx: unknown) => Promise<unknown>) =>
        cb({
          sync: vi.fn(async () => {}),
          workbook: {
            worksheets: {
              items: worksheets,
              load: vi.fn(),
              getActiveWorksheet: () => worksheets[0],
            },
          },
        }),
    };
    return { worksheets, written };
  }

  const writeOp = (range: string) =>
    ({ kind: 'write_range', range, cells: [[{ value: 999 }]], reason: 'test' }) as const;

  it('rejects a write to a missing sheet instead of redirecting it to the active sheet', async () => {
    const { applyOps, readRange } = await import('../src/officeBridge.js');
    const { worksheets, written } = installExcel(['Payroll']);

    await expect(applyOps([writeOp('MissingSheet!A1')])).rejects.toThrow('No such sheet: MissingSheet');
    await expect(readRange('Sheet1!A1')).rejects.toThrow('No such sheet: Sheet1');
    await expect(applyOps([{ kind: 'delete_sheet', name: 'Gone', reason: 'x' }])).rejects.toThrow(
      'No such sheet: Gone',
    );
    expect(worksheets[0]!.getRange).not.toHaveBeenCalled();
    expect(written.size).toBe(0);
  });

  it('rejects unqualified references and adding a sheet that already exists', async () => {
    const { applyOps } = await import('../src/officeBridge.js');
    installExcel(['Payroll']);

    await expect(applyOps([writeOp('A1')])).rejects.toThrow('Reference is missing a sheet name');
    await expect(applyOps([{ kind: 'add_sheet', name: 'payroll', reason: 'x' }])).rejects.toThrow(
      'Sheet already exists: Payroll',
    );
  });

  it('matches sheet names case-insensitively and writes to exactly that sheet', async () => {
    const { applyOps } = await import('../src/officeBridge.js');
    const { written } = installExcel(['Payroll', '2025 Report']);

    await applyOps([writeOp("'2025 report'!A1")]);
    expect([...written.keys()]).toEqual(['2025 Report!A1']);
  });
});
