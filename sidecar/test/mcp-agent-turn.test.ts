import { describe, expect, it, beforeEach, afterEach } from 'vitest';
import { DatabaseSync } from 'node:sqlite';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { WorkbookEmulator } from '../src/workbook/emulator.js';
import { ChangeSetManager } from '../src/changeset/manager.js';
import { ScriptedProvider } from '../src/providers/scripted.js';
import { runAgentTurn, type AgentEvent } from '../src/agent/loop.js';
import { McpManager } from '../src/mcp/manager.js';

describe('agent loop with MCP connectors', () => {
  let tempDir: string;
  let dbPath: string;
  let manager: McpManager;

  beforeEach(async () => {
    tempDir = join(tmpdir(), `excelai-mcp-agent-test-${Date.now()}`);
    mkdirSync(tempDir, { recursive: true });
    dbPath = join(tempDir, 'products.db');

    const db = new DatabaseSync(dbPath);
    db.exec(`
      CREATE TABLE products (id INT, name TEXT, price REAL);
      INSERT INTO products VALUES (1, 'Widget', 25.50);
      INSERT INTO products VALUES (2, 'Gadget', 42.00);
    `);
    db.close();

    manager = new McpManager();
    await manager.sync(
      [
        {
          id: 'prod-db',
          name: 'ProductsDB',
          enabled: true,
          transport: 'builtin',
          type: 'sqlite',
          allowedPaths: [tempDir],
          options: { dbPath },
        },
      ],
      [tempDir],
    );
  });

  afterEach(async () => {
    await manager.closeAll();
    try {
      rmSync(tempDir, { recursive: true, force: true });
    } catch {}
  });

  it('allows model to query SQLite via MCP and stage write_range based on data', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const changeSets = new ChangeSetManager(wb, 'wb-test');
    const events: AgentEvent[] = [];

    // Model queries sqlite_read_query, then writes total to Sheet1!A1
    const provider = new ScriptedProvider([
      {
        toolCalls: [
          {
            name: 'sqlite_read_query',
            args: { query: 'SELECT name, price FROM products' },
          },
        ],
      },
      {
        toolCalls: [
          {
            name: 'write_range',
            args: {
              range: 'Sheet1!A1:B2',
              cells: [
                ['Widget', 25.50],
                ['Gadget', 42.00],
              ],
              reason: 'Imported products from SQLite',
            },
          },
        ],
      },
      { text: 'I fetched products from SQLite and staged them into Sheet1.' },
    ]);

    const result = await runAgentTurn({
      provider,
      executor: wb,
      changeSets,
      workbookName: 'Products.xlsx',
      history: [],
      userMessage: 'Import products from the database',
      mcpManager: manager,
      onEvent: (ev) => events.push(ev),
    });

    expect(result.stagedChangeSetId).toBeDefined();
    const toolUses = events.filter((e) => e.kind === 'tool_use');
    expect(toolUses.map((t: any) => t.tool)).toContain('sqlite_read_query');
    expect(toolUses.map((t: any) => t.tool)).toContain('write_range');

    const staged = events.find((e) => e.kind === 'changeset_staged');
    expect(staged).toBeDefined();
  });

  it('stages change ops generated directly by MCP tool (sqlite_stage_table_to_sheet)', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const changeSets = new ChangeSetManager(wb, 'wb-test');
    const events: AgentEvent[] = [];

    const provider = new ScriptedProvider([
      {
        toolCalls: [
          {
            name: 'sqlite_stage_table_to_sheet',
            args: {
              queryOrTable: 'products',
              targetSheet: 'Sheet1',
              startCell: 'A1',
              hasHeaders: true,
              reason: 'Direct MCP staging of products table',
            },
          },
        ],
      },
      { text: 'Staged the products table into Sheet1.' },
    ]);

    const result = await runAgentTurn({
      provider,
      executor: wb,
      changeSets,
      workbookName: 'Products.xlsx',
      history: [],
      userMessage: 'Load products table into Sheet1',
      mcpManager: manager,
      onEvent: (ev) => events.push(ev),
    });

    expect(result.stagedChangeSetId).toBeDefined();
    const stagedEvent = events.find((e) => e.kind === 'changeset_staged');
    expect(stagedEvent).toBeDefined();
    if (stagedEvent && stagedEvent.kind === 'changeset_staged') {
      expect((stagedEvent.preview[0].op as any).range).toBe('Sheet1!A1:C3');
    }
  });
});
