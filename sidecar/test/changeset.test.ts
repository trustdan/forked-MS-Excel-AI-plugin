import { describe, expect, it } from 'vitest';
import { WorkbookEmulator } from '../src/workbook/emulator.js';
import { ChangeSetManager } from '../src/changeset/manager.js';

describe('change-set lifecycle', () => {
  it('stages, previews, applies and undoes a write', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    wb.setCell('Sheet1!A1', 'old');
    const mgr = new ChangeSetManager(wb, 'wb-1');

    const cs = await mgr.stage([
      { kind: 'write_range', range: 'Sheet1!A1', cells: [[{ value: 'new' }]], reason: 'test' },
    ]);
    expect(cs.status).toBe('staged');
    expect(cs.undoFidelity).toBe('full');

    const preview = await mgr.preview(cs.id);
    expect(preview[0]?.before?.[0]?.[0]?.value).toBe('old');
    expect(wb.getValue('Sheet1!A1')).toBe('old'); // staging never touches the sheet

    const result = await mgr.apply(cs.id);
    expect(result.ok).toBe(true);
    expect(wb.getValue('Sheet1!A1')).toBe('new');

    await mgr.undo(cs.id);
    expect(wb.getValue('Sheet1!A1')).toBe('old');
    expect(mgr.get(cs.id)?.status).toBe('undone');
  });

  it('restores formulas (not just values) on undo', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    wb.setCell('Sheet1!A1', 2);
    wb.setCell('Sheet1!B1', { formula: '=A1*10' });
    const mgr = new ChangeSetManager(wb, 'wb-1');

    const cs = await mgr.stage([
      { kind: 'write_range', range: 'Sheet1!B1', cells: [[{ value: 999 }]], reason: 'flatten' },
    ]);
    await mgr.apply(cs.id);
    expect(wb.getValue('Sheet1!B1')).toBe(999);

    await mgr.undo(cs.id);
    expect(wb.getValue('Sheet1!B1')).toBe(20);
    wb.setCell('Sheet1!A1', 5);
    expect(wb.getValue('Sheet1!B1')).toBe(50); // formula restored, still live
  });

  it('rolls back automatically when a change introduces formula errors', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    wb.setCell('Sheet1!A1', 10);
    const mgr = new ChangeSetManager(wb, 'wb-1');

    const cs = await mgr.stage([
      {
        kind: 'write_range',
        range: 'Sheet1!B1',
        cells: [[{ formula: '=A1/0' }]],
        reason: 'bad formula',
      },
    ]);
    const result = await mgr.apply(cs.id);
    expect(result.ok).toBe(false);
    expect(result.problems[0]).toContain('#DIV/0!');
    expect(wb.getValue('Sheet1!B1')).toBeNull(); // rolled back
    expect(mgr.get(cs.id)?.status).toBe('staged'); // back to staged for iteration
  });

  it('marks change-sets with structural ops as partial undo fidelity', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const mgr = new ChangeSetManager(wb, 'wb-1');
    const cs = await mgr.stage([{ kind: 'add_sheet', name: 'Report', reason: 'new output sheet' }]);
    expect(cs.undoFidelity).toBe('partial');
  });

  it('rejects mismatched cell matrix dimensions', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const mgr = new ChangeSetManager(wb, 'wb-1');
    const cs = await mgr.stage([
      {
        kind: 'write_range',
        range: 'Sheet1!A1:B2',
        cells: [[{ value: 1 }]], // 1x1 into 2x2
        reason: 'bad dims',
      },
    ]);
    await expect(mgr.apply(cs.id)).rejects.toThrow(/does not match range/);
  });

  it('refuses to double-apply or undo a non-applied set', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const mgr = new ChangeSetManager(wb, 'wb-1');
    const cs = await mgr.stage([
      { kind: 'write_range', range: 'Sheet1!A1', cells: [[{ value: 1 }]], reason: 'x' },
    ]);
    await expect(mgr.undo(cs.id)).rejects.toThrow();
    await mgr.apply(cs.id);
    await expect(mgr.apply(cs.id)).rejects.toThrow();
  });
});

describe('exact sheet targeting (F1)', () => {
  const write = (range: string, value: number) =>
    ({ kind: 'write_range', range, cells: [[{ value }]], reason: 'test' }) as const;

  it('resolves sheet names case-insensitively to the exact worksheet at staging', async () => {
    const wb = new WorkbookEmulator(['Payroll']);
    const mgr = new ChangeSetManager(wb, 'wb-1');
    const cs = await mgr.stage([
      write('payroll!A1', 1),
      { kind: 'insert_rows', sheet: 'PAYROLL', at: 0, count: 1, reason: 'test' },
    ]);
    expect(cs.ops[0]).toMatchObject({ range: 'Payroll!A1' });
    expect(cs.ops[1]).toMatchObject({ sheet: 'Payroll' });
    const map = await wb.getWorkbookMap();
    expect(cs.targets).toEqual([{ id: map.sheets[0]!.id, name: 'Payroll' }]);
  });

  it('refuses to stage a write to a sheet that does not exist', async () => {
    const wb = new WorkbookEmulator(['Payroll']);
    const mgr = new ChangeSetManager(wb, 'wb-1');
    await expect(mgr.stage([write('MissingSheet!A1', 999)])).rejects.toThrow(
      /No such sheet: MissingSheet\. Sheets in this workbook: Payroll/,
    );
    expect(mgr.list()).toHaveLength(0);
    expect(wb.getValue('Payroll!A1')).toBeNull();
  });

  it('lets later ops target sheets the same change-set adds or renames', async () => {
    const wb = new WorkbookEmulator(['Data']);
    const mgr = new ChangeSetManager(wb, 'wb-1');
    const cs = await mgr.stage([
      { kind: 'add_sheet', name: 'Report', reason: 'new' },
      write('report!A1', 1),
      { kind: 'rename_sheet', name: 'data', newName: 'Raw', reason: 'rename' },
      write('Raw!B2', 2),
    ]);
    expect(cs.ops.map((op) => ('range' in op ? op.range : op.name))).toEqual([
      'Report',
      'Report!A1',
      'Data',
      'Raw!B2',
    ]);
    // Only the pre-existing sheet is a target, under its name at staging.
    expect(cs.targets.map((t) => t.name)).toEqual(['Data']);
    expect((await mgr.apply(cs.id)).ok).toBe(true);
    expect(wb.getValue('Report!A1')).toBe(1);
    expect(wb.getValue('Raw!B2')).toBe(2);
  });

  it('rejects adding a sheet that already exists, or using a sheet after deleting it', async () => {
    const wb = new WorkbookEmulator(['Data', 'Old']);
    const mgr = new ChangeSetManager(wb, 'wb-1');
    await expect(mgr.stage([{ kind: 'add_sheet', name: 'data', reason: 'x' }])).rejects.toThrow(
      /Sheet already exists: Data/,
    );
    await expect(
      mgr.stage([{ kind: 'delete_sheet', name: 'Old', reason: 'x' }, write('Old!A1', 1)]),
    ).rejects.toThrow(/No such sheet: Old/);
  });

  it('fails approval when the reviewed target sheet was deleted, leaving other sheets untouched', async () => {
    const wb = new WorkbookEmulator(['Payroll', 'Report']);
    const mgr = new ChangeSetManager(wb, 'wb-1');
    const cs = await mgr.stage([write('Report!A1', 999)]);

    await wb.applyOps([{ kind: 'delete_sheet', name: 'Report', reason: 'user deleted it' }]);

    await expect(mgr.apply(cs.id)).rejects.toThrow(/Sheet "Report" no longer exists/);
    expect(wb.getValue('Payroll!A1')).toBeNull();
    expect(mgr.get(cs.id)?.status).toBe('staged');
  });

  it('fails approval explicitly when the target was renamed, even if another sheet took its name', async () => {
    const wb = new WorkbookEmulator(['Payroll', 'Report']);
    const mgr = new ChangeSetManager(wb, 'wb-1');
    const cs = await mgr.stage([write('Report!A1', 999)]);

    await wb.applyOps([
      { kind: 'rename_sheet', name: 'Report', newName: 'Report 2025', reason: 'user' },
      { kind: 'rename_sheet', name: 'Payroll', newName: 'Report', reason: 'user' },
    ]);

    await expect(mgr.apply(cs.id)).rejects.toThrow(
      /Sheet "Report" was renamed to "Report 2025"/,
    );
    expect(wb.getValue('Report!A1')).toBeNull(); // the former Payroll sheet
    expect(wb.getValue("'Report 2025'!A1")).toBeNull();
  });

  it('keeps emulator sheet ids stable across renames', async () => {
    const wb = new WorkbookEmulator(['A', 'B']);
    const [a, b] = (await wb.getWorkbookMap()).sheets;
    expect(a!.id).not.toBe(b!.id);
    await wb.applyOps([{ kind: 'rename_sheet', name: 'A', newName: 'C', reason: 'x' }]);
    const after = (await wb.getWorkbookMap()).sheets;
    expect(after.find((s) => s.name === 'C')?.id).toBe(a!.id);
  });
});
