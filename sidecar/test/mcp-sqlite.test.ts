import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { DatabaseSync } from 'node:sqlite';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { mkdirSync, rmSync } from 'node:fs';
import { SqliteConnector } from '../src/mcp/connectors/sqliteConnector.js';
import { McpSecurityError } from '../src/mcp/pathGuard.js';
import { openDb } from '../src/db.js';
import { startServer } from '../src/server.js';

const ROWS_1500 = 'WITH RECURSIVE n(x) AS (VALUES(1) UNION ALL SELECT x+1 FROM n WHERE x<1500)';
// About six seconds of single-step work: long enough that a blocked event loop is unmistakable.
const SLOW_QUERY =
  'WITH RECURSIVE n(x) AS (VALUES(1) UNION ALL SELECT x+1 FROM n WHERE x<60000000) SELECT count(*) AS c FROM n';

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

  describe('bounded work (F10)', () => {
    const open = (options: Record<string, unknown> = {}): SqliteConnector =>
      (connector = new SqliteConnector({
        id: 'sqlite-1',
        name: 'Sales DB',
        enabled: true,
        transport: 'builtin',
        type: 'sqlite',
        allowedPaths: [tempDir],
        options: { dbPath, ...options },
      }));
    const rowsOf = (res: { content: { text: string }[] }): { rowCount: number; truncated: boolean } =>
      JSON.parse(res.content[0]!.text);

    it.each([
      ['an explicit larger LIMIT', `${ROWS_1500} SELECT x FROM n LIMIT 1500`],
      ['LIMIT -1', `${ROWS_1500} SELECT x FROM n LIMIT -1`],
      ['a LIMIT inside a comment', `${ROWS_1500} SELECT x FROM n -- limit`],
      ['a LIMIT inside a string literal', `${ROWS_1500} SELECT x, 'limit' AS l FROM n`],
      ['a nested LIMIT', `${ROWS_1500} SELECT x FROM (SELECT x FROM n LIMIT 1500)`],
    ])('caps rows at the requested limit despite %s', async (_label, query) => {
      const res = await open().callTool('sqlite_read_query', { query, limit: 1 });
      expect(res.isError).toBeFalsy();
      expect(rowsOf(res)).toMatchObject({ rowCount: 1, truncated: true });
    });

    it('caps staged rows at 1000 whatever the SQL says', async () => {
      const res = await open().callTool('sqlite_stage_table_to_sheet', {
        queryOrTable: `${ROWS_1500} SELECT x FROM n LIMIT 1500`,
        hasHeaders: true,
      });
      expect(res.isError).toBeFalsy();
      const op = res.changeOps?.[0];
      expect(op?.kind === 'write_range' && op.cells.length).toBe(1001);
      expect(JSON.parse(res.content[0]!.text)).toMatchObject({ rowCount: 1000, truncated: true });
    });

    it('keeps /api/health responsive during a slow query and stops it on abort', async () => {
      const sidecar = await startServer({ db: openDb(':memory:'), port: 0 });
      try {
        const controller = new AbortController();
        const pending = open().callTool('sqlite_read_query', { query: SLOW_QUERY }, {
          abortSignal: controller.signal,
        });
        const settled = pending.then(
          () => 'resolved',
          (e: Error) => e.message,
        );
        await new Promise((r) => setTimeout(r, 300));

        const t0 = Date.now();
        const health = await fetch(`http://127.0.0.1:${sidecar.port}/api/health`);
        expect(health.status).toBe(200);
        expect(Date.now() - t0).toBeLessThan(1_000);

        const t1 = Date.now();
        controller.abort();
        expect(await settled).toMatch(/stopped/i);
        expect(Date.now() - t1).toBeLessThan(1_000);
      } finally {
        await sidecar.close();
      }
    });

    it('stops a query that exceeds its execution budget', async () => {
      const t0 = Date.now();
      await expect(
        open({ queryTimeoutMs: 500 }).callTool('sqlite_read_query', { query: SLOW_QUERY }),
      ).rejects.toThrow(/execution budget/);
      expect(Date.now() - t0).toBeLessThan(3_000);
    });

    it('rejects malformed limits instead of coercing them', async () => {
      const c = open();
      await expect(c.callTool('sqlite_read_query', { query: 'SELECT 1', limit: 'invalid' })).rejects.toThrow(
        /limit/,
      );
      await expect(c.callTool('sqlite_read_query', { query: 'SELECT 1', limit: 5000 })).rejects.toThrow(/limit/);
      await expect(c.callTool('sqlite_read_query', { query: 'SELECT 1', limit: 2.5 })).rejects.toThrow(/limit/);
    });
  });
});
