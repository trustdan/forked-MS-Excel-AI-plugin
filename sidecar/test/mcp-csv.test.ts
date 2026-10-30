import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { mkdirSync, writeFileSync, rmSync } from 'node:fs';
import { CsvDirectoryConnector } from '../src/mcp/connectors/csvDirectoryConnector.js';
import { McpSecurityError } from '../src/mcp/pathGuard.js';

describe('CsvDirectoryConnector', () => {
  let tempDir: string;
  let csvDir: string;
  let connector: CsvDirectoryConnector | null = null;

  beforeEach(() => {
    tempDir = join(tmpdir(), `excelai-csv-test-${Date.now()}`);
    csvDir = join(tempDir, 'data');
    mkdirSync(csvDir, { recursive: true });

    writeFileSync(
      join(csvDir, 'employees.csv'),
      'id,name,department,salary\n1,Alice,Engineering,120000\n2,Bob,Marketing,95000\n3,Charlie,Engineering,110000\n',
    );
    writeFileSync(join(csvDir, 'notes.txt'), 'Not a CSV file');

    const sub = join(csvDir, 'sub');
    mkdirSync(sub, { recursive: true });
    writeFileSync(join(sub, 'archive.tsv'), 'month\trevenue\nJan\t1000\nFeb\t1500\n');
  });

  afterEach(async () => {
    if (connector) {
      await connector.close();
      connector = null;
    }
    try {
      rmSync(tempDir, { recursive: true, force: true });
    } catch {}
  });

  it('rejects directory outside approved paths', async () => {
    connector = new CsvDirectoryConnector({
      id: 'csv-1',
      name: 'CSV Store',
      enabled: true,
      transport: 'builtin',
      type: 'csv-directory',
      allowedPaths: [join(tempDir, 'non-approved')],
      options: { dirPath: csvDir },
    });

    await expect(connector.initialize()).rejects.toThrow(McpSecurityError);
  });

  it('lists CSV and TSV files in directory tree', async () => {
    connector = new CsvDirectoryConnector({
      id: 'csv-1',
      name: 'CSV Store',
      enabled: true,
      transport: 'builtin',
      type: 'csv-directory',
      allowedPaths: [csvDir],
      options: { dirPath: csvDir },
    });

    await connector.initialize();
    const res = await connector.callTool('csv_list_files', {});
    expect(res.isError).toBeFalsy();
    const data = JSON.parse(res.content[0].text);
    expect(data.ok).toBe(true);
    expect(data.totalFiles).toBe(2);
    const paths = data.files.map((f: any) => f.relativePath);
    expect(paths).toContain('employees.csv');
    expect(paths).toContain('sub/archive.tsv');
  });

  it('reads preview of CSV file', async () => {
    connector = new CsvDirectoryConnector({
      id: 'csv-1',
      name: 'CSV Store',
      enabled: true,
      transport: 'builtin',
      type: 'csv-directory',
      allowedPaths: [csvDir],
      options: { dirPath: csvDir },
    });

    const res = await connector.callTool('csv_read_preview', {
      relativePath: 'employees.csv',
      maxRows: 2,
    });
    expect(res.isError).toBeFalsy();
    const data = JSON.parse(res.content[0].text);
    expect(data.ok).toBe(true);
    expect(data.previewRowCount).toBe(2);
    expect(data.rows[0]).toEqual(['id', 'name', 'department', 'salary']);
    expect(data.rows[1]).toEqual([1, 'Alice', 'Engineering', 120000]);
  });

  it('queries and filters CSV rows by column value', async () => {
    connector = new CsvDirectoryConnector({
      id: 'csv-1',
      name: 'CSV Store',
      enabled: true,
      transport: 'builtin',
      type: 'csv-directory',
      allowedPaths: [csvDir],
      options: { dirPath: csvDir },
    });

    const res = await connector.callTool('csv_query', {
      relativePath: 'employees.csv',
      column: 'department',
      filterValue: 'Engineering',
    });
    expect(res.isError).toBeFalsy();
    const data = JSON.parse(res.content[0].text);
    expect(data.ok).toBe(true);
    expect(data.totalMatched).toBe(2); // Alice and Charlie
    expect(data.rows[1][1]).toBe('Alice');
    expect(data.rows[2][1]).toBe('Charlie');
  });

  it('stages CSV directly into Excel worksheet change-set', async () => {
    connector = new CsvDirectoryConnector({
      id: 'csv-1',
      name: 'CSV Store',
      enabled: true,
      transport: 'builtin',
      type: 'csv-directory',
      allowedPaths: [csvDir],
      options: { dirPath: csvDir },
    });

    const res = await connector.callTool('csv_stage_to_sheet', {
      relativePath: 'employees.csv',
      targetSheet: 'Employees',
      startCell: 'B2',
      hasHeaders: true,
      reason: 'Stage employee payroll data',
    });

    expect(res.isError).toBeFalsy();
    expect(res.changeOps).toHaveLength(2); // write_range + header format_range
    const writeOp = res.changeOps?.find((op) => op.kind === 'write_range');
    expect(writeOp).toBeDefined();
    if (writeOp && writeOp.kind === 'write_range') {
      expect(writeOp.range).toBe('Employees!B2:E5');
    }
  });

  it('rejects path traversal attempting to escape CSV directory', async () => {
    connector = new CsvDirectoryConnector({
      id: 'csv-1',
      name: 'CSV Store',
      enabled: true,
      transport: 'builtin',
      type: 'csv-directory',
      allowedPaths: [csvDir],
      options: { dirPath: csvDir },
    });

    await expect(
      connector.callTool('csv_read_preview', {
        relativePath: '../secret.csv',
      }),
    ).rejects.toThrow(McpSecurityError);
  });

  it('rejects malformed numeric arguments instead of coercing them (F11)', async () => {
    connector = new CsvDirectoryConnector({
      id: 'csv-1',
      name: 'CSV Data',
      enabled: true,
      transport: 'builtin',
      type: 'csv-directory',
      allowedPaths: [csvDir],
      options: { dirPath: csvDir },
    });

    await expect(
      connector.callTool('csv_read_preview', { relativePath: 'employees.csv', maxRows: 'invalid' }),
    ).rejects.toThrow(/maxRows/);
    await expect(
      connector.callTool('csv_query', { relativePath: 'employees.csv', limit: 10_000 }),
    ).rejects.toThrow(/limit/);
  });
});
