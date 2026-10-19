import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { DatabaseSync } from 'node:sqlite';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { mkdirSync, rmSync } from 'node:fs';
import { SqliteConnector } from '../src/mcp/connectors/sqliteConnector.js';
import { McpSecurityError } from '../src/mcp/pathGuard.js';

describe('SqliteConnector', () => {
  let tempDir: string;
  let dbPath: string;
  let connector: SqliteConnector | null = null;

  beforeEach(() => {
    tempDir = join(tmpdir(), `excelai-sqlite-test-${Date.now()}`);
    mkdirSync(tempDir, { recursive: true });
    dbPath = join(tempDir, 'sales.db');

    // Create a real sqlite db with tables
    const db = new DatabaseSync(dbPath);
    db.exec(`
      CREATE TABLE customers (
        id INTEGER PRIMARY KEY,
        name TEXT NOT NULL,
        balance REAL
      );
      INSERT INTO customers VALUES (1, 'Alice', 150.50);
      INSERT INTO customers VALUES (2, 'Bob', 200.00);
      INSERT INTO customers VALUES (3, '=cmd|'' /C calc''!A0', 99.99);

      CREATE TABLE orders (
        order_id INTEGER PRIMARY KEY,
        customer_id INTEGER,
        amount REAL
      );
      INSERT INTO orders VALUES (101, 1, 50.00);
    `);
    db.close();
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

  it('rejects database path outside approved paths', async () => {
    connector = new SqliteConnector({
      id: 'sqlite-1',
      name: 'Sales DB',
      enabled: true,
      transport: 'builtin',
      type: 'sqlite',
      allowedPaths: [join(tempDir, 'other-folder')],
      options: { dbPath },
    });

    await expect(connector.initialize()).rejects.toThrow(McpSecurityError);
  });

  it('lists user tables with row counts and tool definitions', async () => {
    connector = new SqliteConnector({
      id: 'sqlite-1',
      name: 'Sales DB',
      enabled: true,
      transport: 'builtin',
      type: 'sqlite',
      allowedPaths: [tempDir],
      options: { dbPath },
    });

    await connector.initialize();
    const tools = await connector.listTools();
    expect(tools.map((t) => t.name)).toContain('sqlite_list_tables');
    expect(tools.map((t) => t.name)).toContain('sqlite_describe_table');
    expect(tools.map((t) => t.name)).toContain('sqlite_read_query');
    expect(tools.map((t) => t.name)).toContain('sqlite_stage_table_to_sheet');

    const res = await connector.callTool('sqlite_list_tables', {});
    expect(res.isError).toBeFalsy();
    const data = JSON.parse(res.content[0].text);
    expect(data.ok).toBe(true);
    expect(data.tables).toHaveLength(2);
    expect(data.tables.find((t: any) => t.name === 'customers').rowCount).toBe(3);
  });

  it('describes table schema with columns, types and primary keys', async () => {
    connector = new SqliteConnector({
      id: 'sqlite-1',
      name: 'Sales DB',
      enabled: true,
      transport: 'builtin',
      type: 'sqlite',
      allowedPaths: [tempDir],
      options: { dbPath },
    });

    const res = await connector.callTool('sqlite_describe_table', { table: 'customers' });
    expect(res.isError).toBeFalsy();
    const data = JSON.parse(res.content[0].text);
    expect(data.ok).toBe(true);
    expect(data.columns).toHaveLength(3);
    const idCol = data.columns.find((c: any) => c.name === 'id');
    expect(idCol.primaryKey).toBe(true);
  });

  it('executes read-only SELECT queries', async () => {
    connector = new SqliteConnector({
      id: 'sqlite-1',
      name: 'Sales DB',
      enabled: true,
      transport: 'builtin',
      type: 'sqlite',
      allowedPaths: [tempDir],
      options: { dbPath },
    });

    const res = await connector.callTool('sqlite_read_query', {
      query: 'SELECT name, balance FROM customers WHERE balance > 100 ORDER BY balance DESC',
    });
    expect(res.isError).toBeFalsy();
    const data = JSON.parse(res.content[0].text);
    expect(data.rowCount).toBe(2);
    expect(data.rows[0].name).toBe('Bob');
  });

  it('rejects mutating SQL statements (INSERT / UPDATE / DROP / DELETE)', async () => {
    connector = new SqliteConnector({
      id: 'sqlite-1',
      name: 'Sales DB',
      enabled: true,
      transport: 'builtin',
      type: 'sqlite',
      allowedPaths: [tempDir],
      options: { dbPath },
    });

    await expect(
      connector.callTool('sqlite_read_query', {
        query: 'DELETE FROM customers WHERE id = 1',
      }),
    ).rejects.toThrow(McpSecurityError);

    await expect(
      connector.callTool('sqlite_read_query', {
        query: 'DROP TABLE customers',
      }),
    ).rejects.toThrow(McpSecurityError);
  });

  it('stages table data directly to sheet with formula escaping', async () => {
    connector = new SqliteConnector({
      id: 'sqlite-1',
      name: 'Sales DB',
      enabled: true,
      transport: 'builtin',
      type: 'sqlite',
      allowedPaths: [tempDir],
      options: { dbPath },
    });

    const res = await connector.callTool('sqlite_stage_table_to_sheet', {
      queryOrTable: 'customers',
      targetSheet: 'Sheet1',
      startCell: 'A1',
      hasHeaders: true,
      reason: 'Import customer records',
    });

    expect(res.isError).toBeFalsy();
    expect(res.changeOps).toHaveLength(1);
    const op = res.changeOps![0];
    expect(op.kind).toBe('write_range');
    if (op.kind === 'write_range') {
      expect(op.range).toBe('Sheet1!A1:C4'); // 1 header row + 3 data rows, 3 cols
      // Header row
      expect(op.cells[0][0]).toEqual({ value: 'id' });
      expect(op.cells[0][1]).toEqual({ value: 'name' });
      // Row with malicious formula should be escaped
      const maliciousName = op.cells[3][1];
      expect(maliciousName).toEqual({ value: "'=cmd|' /C calc'!A0" });
    }
  });
});
