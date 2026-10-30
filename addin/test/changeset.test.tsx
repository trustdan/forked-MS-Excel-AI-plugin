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

  it('applyOps and readRange fall back to activeWorksheet instead of throwing ItemNotFound ("The requested resource doesn\'t exist.") when Sheet1 is missing', async () => {
    const { applyOps, readRange } = await import('../src/officeBridge.js');

    let writtenFormulas: unknown[][] | null = null;
    const activeRange = {
      values: [['Cash', 'Raw Materials'], [401900, 419500]],
      formulas: [['Cash', 'Raw Materials'], [401900, 419500]],
      load: vi.fn(),
      get setFormulas() {
        return writtenFormulas;
      },
      set formulas_setter(v: unknown[][]) {
        writtenFormulas = v;
      },
    };
    Object.defineProperty(activeRange, 'formulas', {
      get: () => writtenFormulas ?? [['Cash', 'Raw Materials'], [401900, 419500]],
      set: (v: unknown[][]) => {
        writtenFormulas = v;
      },
    });

    const activeWorksheet = {
      name: '2025 Starting Balances',
      getRange: vi.fn(() => activeRange),
    };

    (globalThis as unknown as { Excel: unknown }).Excel = {
      run: async (cb: (ctx: unknown) => Promise<unknown>) => {
        const ctx = {
          sync: vi.fn(async () => {}),
          workbook: {
            worksheets: {
              items: [activeWorksheet],
              load: vi.fn(),
              getActiveWorksheet: () => activeWorksheet,
              getItem: (name: string) => {
                if (name !== '2025 Starting Balances') {
                  throw new Error("The requested resource doesn't exist.");
                }
                return activeWorksheet;
              },
            },
          },
        };
        return cb(ctx);
      },
    };

    await applyOps([
      {
        kind: 'write_range',
        range: 'Sheet1!A1:B2',
        cells: [
          [{ value: 'Cash' }, { value: 'Raw Materials' }],
          [{ value: 401900 }, { value: 419500 }],
        ],
        reason: 'List 2025 starting balances horizontally',
      },
    ]);

    expect(activeWorksheet.getRange).toHaveBeenCalledWith('A1:B2');
    expect(writtenFormulas).toEqual([
      ['Cash', 'Raw Materials'],
      [401900, 419500],
    ]);

    const readBack = await readRange('Sheet1!A1:B2');
    expect(readBack[0]?.[0]?.value).toBe('Cash');
    expect(readBack[1]?.[0]?.value).toBe(401900);
  });
});
