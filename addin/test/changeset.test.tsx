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
});
