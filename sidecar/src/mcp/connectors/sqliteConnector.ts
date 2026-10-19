import { DatabaseSync } from 'node:sqlite';
import { existsSync } from 'node:fs';
import type {
  ChangeOp,
  McpServerConfig,
  McpServerStatus,
  McpToolDefinition,
  McpToolResult,
} from '@excelai/shared';
import { indexToCol, parseA1, formatA1 } from '@excelai/shared';
import { PathGuard, McpSecurityError } from '../pathGuard.js';
import type { McpClient } from '../types.js';

const DISALLOWED_SQL_PATTERNS = [
  /\b(insert|update|delete|drop|alter|create|replace|truncate|vacuum|attach|detach)\b/i,
  /\bpragma\s+(?!table_info|table_xinfo)/i,
];

export class SqliteConnector implements McpClient {
  readonly id: string;
  readonly name: string;
  readonly config: McpServerConfig;
  private db: DatabaseSync | null = null;
  private dbPath: string | null = null;
  private statusState: 'connected' | 'disconnected' | 'error' = 'disconnected';
  private lastError?: string;

  constructor(config: McpServerConfig) {
    this.id = config.id;
    this.name = config.name;
    this.config = config;
  }

  async initialize(): Promise<void> {
    try {
      const rawDbPath = String(this.config.options?.dbPath ?? '');
      if (!rawDbPath) {
        throw new Error('options.dbPath is required for SQLite connector');
      }
      this.dbPath = PathGuard.assertApproved(
        rawDbPath,
        this.config.allowedPaths,
        `SQLite Connector [${this.name}]`,
      );

      if (!existsSync(this.dbPath)) {
        throw new Error(`Database file does not exist: ${this.dbPath}`);
      }

      this.db = new DatabaseSync(this.dbPath, { readOnly: true, open: true });
      this.statusState = 'connected';
      this.lastError = undefined;
    } catch (e) {
      this.statusState = 'error';
      this.lastError = e instanceof Error ? e.message : String(e);
      throw e;
    }
  }

  async listTools(): Promise<McpToolDefinition[]> {
    return [
      {
        name: 'sqlite_list_tables',
        description: 'List all user tables and views in the connected SQLite database with column counts.',
        inputSchema: {
          type: 'object',
          properties: {},
          additionalProperties: false,
        },
      },
      {
        name: 'sqlite_describe_table',
        description: 'Describe the schema of a table or view (columns, data types, primary keys, nullability).',
        inputSchema: {
          type: 'object',
          properties: {
            table: {
              type: 'string',
              description: 'The table or view name to inspect',
            },
          },
          required: ['table'],
          additionalProperties: false,
        },
      },
      {
        name: 'sqlite_read_query',
        description:
          'Execute a read-only SQL query (SELECT / WITH / EXPLAIN) against the SQLite database. Mutation queries are strictly forbidden.',
        inputSchema: {
          type: 'object',
          properties: {
            query: {
              type: 'string',
              description: 'Read-only SQL statement (e.g. "SELECT * FROM orders WHERE total > 100 LIMIT 50")',
            },
            limit: {
              type: 'number',
              description: 'Maximum rows to return (default 500, max 1000)',
            },
          },
          required: ['query'],
          additionalProperties: false,
        },
      },
      {
        name: 'sqlite_stage_table_to_sheet',
        description:
          'Execute a read-only query or select a full table and stage the resulting rows directly into an Excel worksheet change-set for user review.',
        inputSchema: {
          type: 'object',
          properties: {
            queryOrTable: {
              type: 'string',
              description: 'A table name or a SELECT statement',
            },
            targetSheet: {
              type: 'string',
              description: 'Sheet name to populate (defaults to "Sheet1")',
            },
            startCell: {
              type: 'string',
              description: 'Top-left cell address (defaults to "A1")',
            },
            hasHeaders: {
              type: 'boolean',
              description: 'Whether to include column headers in row 1 (defaults to true)',
            },
            reason: {
              type: 'string',
              description: 'One-line justification for staging this data',
            },
          },
          required: ['queryOrTable'],
          additionalProperties: false,
        },
      },
    ];
  }

  async callTool(
    toolName: string,
    args: Record<string, unknown>,
    _options?: { abortSignal?: AbortSignal },
  ): Promise<McpToolResult> {
    if (!this.db) {
      await this.initialize();
    }
    const db = this.db!;

    switch (toolName) {
      case 'sqlite_list_tables': {
        const stmt = db.prepare(
          "SELECT name, type FROM sqlite_master WHERE type IN ('table', 'view') AND name NOT LIKE 'sqlite_%' ORDER BY name ASC;",
        );
        const rows = stmt.all() as Array<{ name: string; type: string }>;
        const tablesWithCounts = rows.map((r) => {
          let count = 0;
          try {
            // Count rows if it's a table
            if (r.type === 'table') {
              const countStmt = db.prepare(`SELECT count(*) as c FROM "${r.name.replace(/"/g, '""')}";`);
              const res = countStmt.get() as { c: number } | undefined;
              count = res?.c ?? 0;
            }
          } catch {
            // view or error
          }
          return { name: r.name, type: r.type, rowCount: count };
        });
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ ok: true, tables: tablesWithCounts }),
            },
          ],
        };
      }

      case 'sqlite_describe_table': {
        const tableName = String(args.table ?? '').trim();
        if (!tableName) {
          return { content: [{ type: 'text', text: 'Error: table name is required' }], isError: true };
        }
        // Validate table name contains no injection chars
        if (!/^[a-zA-Z0-9_ -]+$/.test(tableName)) {
          return {
            content: [{ type: 'text', text: 'Error: invalid characters in table name' }],
            isError: true,
          };
        }
        const stmt = db.prepare(`PRAGMA table_info("${tableName.replace(/"/g, '""')}");`);
        const columns = stmt.all() as Array<{
          cid: number;
          name: string;
          type: string;
          notnull: number;
          dflt_value: unknown;
          pk: number;
        }>;
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                ok: true,
                table: tableName,
                columns: columns.map((c) => ({
                  cid: c.cid,
                  name: c.name,
                  type: c.type,
                  notNull: c.notnull === 1,
                  defaultValue: c.dflt_value,
                  primaryKey: c.pk > 0,
                })),
              }),
            },
          ],
        };
      }

      case 'sqlite_read_query': {
        const query = String(args.query ?? '').trim();
        const limit = Math.min(Math.max(1, Number(args.limit ?? 500)), 1000);
        this.assertReadOnlyQuery(query);

        let finalQuery = query;
        if (!/\blimit\b/i.test(finalQuery)) {
          finalQuery = `${finalQuery} LIMIT ${limit}`;
        }

        const stmt = db.prepare(finalQuery);
        const rows = stmt.all() as Record<string, unknown>[];
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                ok: true,
                rowCount: rows.length,
                columns: rows.length > 0 && rows[0] ? Object.keys(rows[0]) : [],
                rows,
              }),
            },
          ],
        };
      }

      case 'sqlite_stage_table_to_sheet': {
        const rawTarget = String(args.queryOrTable ?? '').trim();
        const targetSheet = String(args.targetSheet ?? 'Sheet1').trim();
        const startCell = String(args.startCell ?? 'A1').trim();
        const hasHeaders = args.hasHeaders !== false;
        const reason = String(args.reason ?? `Import SQLite data: ${rawTarget.slice(0, 50)}`);

        let query = rawTarget;
        if (!/\bselect\b/i.test(query)) {
          // Treat as table name
          if (!/^[a-zA-Z0-9_ -]+$/.test(rawTarget)) {
            return {
              content: [{ type: 'text', text: 'Error: invalid table name' }],
              isError: true,
            };
          }
          query = `SELECT * FROM "${rawTarget.replace(/"/g, '""')}" LIMIT 1000;`;
        } else {
          this.assertReadOnlyQuery(query);
          if (!/\blimit\b/i.test(query)) {
            query = `${query} LIMIT 1000`;
          }
        }

        const stmt = db.prepare(query);
        const rows = stmt.all() as Record<string, unknown>[];
        if (rows.length === 0 || !rows[0]) {
          return {
            content: [{ type: 'text', text: JSON.stringify({ ok: true, note: 'Query returned 0 rows. No cells staged.' }) }],
          };
        }

        const columns = Object.keys(rows[0]);
        const grid: (string | number | boolean | null)[][] = [];

        if (hasHeaders) {
          grid.push(columns);
        }

        for (const row of rows) {
          const rowValues: (string | number | boolean | null)[] = [];
          for (const col of columns) {
            const val = row[col];
            if (val === null || val === undefined) {
              rowValues.push(null);
            } else if (typeof val === 'number' || typeof val === 'boolean') {
              rowValues.push(val);
            } else if (typeof val === 'string') {
              // Formula injection guard (CWE-1236)
              if (/^[=@+-]/.test(val) && !/^[+-]?\d+(\.\d+)?$/.test(val)) {
                rowValues.push(`'${val}`);
              } else {
                rowValues.push(val);
              }
            } else {
              rowValues.push(String(val));
            }
          }
          grid.push(rowValues);
        }

        const startAddr = parseA1(`${targetSheet}!${startCell}`);
        const totalRows = grid.length;
        const totalCols = columns.length;
        const targetRange = formatA1({
          sheet: targetSheet,
          startRow: startAddr.startRow,
          startCol: startAddr.startCol,
          endRow: startAddr.startRow + totalRows - 1,
          endCol: startAddr.startCol + totalCols - 1,
        });

        const changeOp: ChangeOp = {
          kind: 'write_range',
          range: targetRange,
          cells: grid.map((r) => r.map((c) => ({ value: c }))),
          reason,
        };

        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                ok: true,
                staged: true,
                targetRange,
                rowCount: rows.length,
                columnCount: columns.length,
                note: `Staged ${rows.length} rows into ${targetRange}. Awaiting user review and approval.`,
              }),
            },
          ],
          changeOps: [changeOp],
        };
      }

      default:
        return {
          content: [{ type: 'text', text: `Unknown tool: ${toolName}` }],
          isError: true,
        };
    }
  }

  private assertReadOnlyQuery(query: string): void {
    const trimmed = query.trim();
    if (!trimmed) {
      throw new Error('Query cannot be empty');
    }
    // Must start with SELECT, WITH, or EXPLAIN
    if (!/^(select|with|explain)\b/i.test(trimmed)) {
      throw new McpSecurityError(
        'SecurityError: Only read-only queries (SELECT, WITH, EXPLAIN) are permitted.',
      );
    }
    // Check against disallowed mutation keywords
    for (const pattern of DISALLOWED_SQL_PATTERNS) {
      if (pattern.test(trimmed)) {
        throw new McpSecurityError(
          `SecurityError: Query contains disallowed SQL statement: ${pattern.toString()}`,
        );
      }
    }
  }

  async close(): Promise<void> {
    if (this.db) {
      try {
        this.db.close();
      } catch {
        // ignore already closed
      }
      this.db = null;
    }
    this.statusState = 'disconnected';
  }

  getStatus(): McpServerStatus {
    return {
      id: this.id,
      name: this.name,
      enabled: this.config.enabled,
      transport: 'builtin',
      type: 'sqlite',
      status: this.statusState,
      error: this.lastError,
      toolCount: 4,
    };
  }
}
