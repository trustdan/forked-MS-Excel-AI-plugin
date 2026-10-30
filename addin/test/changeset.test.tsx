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
        status="staged"
        onApprove={() => {}}
        onReject={() => {}}
        onUndo={() => {}}
      />,
    );
    expect(screen.getByText('Sheets: 2025 Payroll, Summary')).toBeDefined();
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
