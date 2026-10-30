import { DatabaseSync } from 'node:sqlite';
import { z } from 'zod';
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
import { parseArgs, toolDefinitions } from '../builtinTools.js';
import { DEFAULT_QUERY_TIMEOUT_MS, runBoundedQuery } from './sqliteQuery.js';

const MAX_QUERY_ROWS = 1000;

const TOOLS = {
  sqlite_list_tables: {
    description: 'List all user tables and views in the connected SQLite database with column counts.',
    parameters: z.object({}).strict(),
  },
  sqlite_describe_table: {
    description: 'Describe the schema of a table or view (columns, data types, primary keys, nullability).',
    parameters: z
      .object({
        table: z.string().trim().min(1).describe('The table or view name to inspect'),
      })
      .strict(),
  },
  sqlite_read_query: {
    description:
      'Execute a read-only SQL query (SELECT / WITH / EXPLAIN) against the SQLite database. Mutation queries are strictly forbidden. At most `limit` rows are returned whatever the SQL says; `truncated` reports whether more existed.',
    parameters: z
      .object({
        query: z
          .string()
          .trim()
          .min(1)
          .describe('Read-only SQL statement (e.g. "SELECT * FROM orders WHERE total > 100")'),
        limit: z
          .number()
          .int()
          .min(1)
          .max(MAX_QUERY_ROWS)
          .optional()
          .describe('Maximum rows to return (default 500, max 1000)'),
      })
      .strict(),
  },
  sqlite_stage_table_to_sheet: {
    description:
      'Execute a read-only query or select a full table and stage the resulting rows (at most 1000) directly into an Excel worksheet change-set for user review.',
    parameters: z
      .object({
        queryOrTable: z.string().trim().min(1).describe('A table name or a SELECT statement'),
        targetSheet: z.string().optional().describe('Sheet name to populate (defaults to "Sheet1")'),
        startCell: z.string().optional().describe('Top-left cell address (defaults to "A1")'),
        hasHeaders: z.boolean().optional().describe('Whether to include column headers in row 1 (defaults to true)'),
        reason: z.string().max(300).optional().describe('One-line justification for staging this data'),
      })
      .strict(),
  },
};

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
    return toolDefinitions(TOOLS);
  }

  async callTool(
    toolName: string,
    args: Record<string, unknown>,
    options?: { abortSignal?: AbortSignal },
  ): Promise<McpToolResult> {
    if (!this.db) {
      await this.initialize();
    }
    const db = this.db!;
    // Anything that scans user data runs out of process under a time budget; schema lookups stay here.
    const query = (sql: string, maxRows: number) =>
      runBoundedQuery({
        dbPath: this.dbPath!,
        sql,
        maxRows,
        timeoutMs: this.queryTimeoutMs(),
        abortSignal: options?.abortSignal,
      });

    switch (toolName) {
      case 'sqlite_list_tables': {
        parseArgs(TOOLS.sqlite_list_tables, args);
        const stmt = db.prepare(
          "SELECT name, type FROM sqlite_master WHERE type IN ('table', 'view') AND name NOT LIKE 'sqlite_%' ORDER BY name ASC;",
        );
        const rows = stmt.all() as Array<{ name: string; type: string }>;
        const tables = rows.filter((r) => r.type === 'table');
        const counts = new Map<string, number>();
        if (tables.length > 0) {
          const countSql = tables
            .map(
              (t) =>
                `SELECT '${t.name.replace(/'/g, "''")}' AS name, count(*) AS c FROM "${t.name.replace(/"/g, '""')}"`,
            )
            .join(' UNION ALL ');
          const res = await query(countSql, tables.length);
          for (const r of res.rows) counts.set(String(r.name), Number(r.c));
        }
        const tablesWithCounts = rows.map((r) => ({ name: r.name, type: r.type, rowCount: counts.get(r.name) ?? 0 }));
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
        const tableName = parseArgs(TOOLS.sqlite_describe_table, args).table;
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
        const { query: sql, limit = 500 } = parseArgs(TOOLS.sqlite_read_query, args);
        this.assertReadOnlyQuery(sql);

        const { rows, truncated } = await query(sql, limit);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                ok: true,
                rowCount: rows.length,
                truncated,
                columns: rows.length > 0 && rows[0] ? Object.keys(rows[0]) : [],
                rows,
              }),
            },
          ],
        };
      }

      case 'sqlite_stage_table_to_sheet': {
        const staged = parseArgs(TOOLS.sqlite_stage_table_to_sheet, args);
        const rawTarget = staged.queryOrTable;
        const targetSheet = (staged.targetSheet ?? 'Sheet1').trim();
        const startCell = (staged.startCell ?? 'A1').trim();
        const hasHeaders = staged.hasHeaders !== false;
        const reason = staged.reason || `Import SQLite data: ${rawTarget.slice(0, 50)}`;

        let sql = rawTarget;
        if (!/\bselect\b/i.test(sql)) {
          // Treat as table name
          if (!/^[a-zA-Z0-9_ -]+$/.test(rawTarget)) {
            return {
              content: [{ type: 'text', text: 'Error: invalid table name' }],
              isError: true,
            };
          }
          sql = `SELECT * FROM "${rawTarget.replace(/"/g, '""')}"`;
        } else {
          this.assertReadOnlyQuery(sql);
        }

        const { rows, truncated } = await query(sql, MAX_QUERY_ROWS);
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
                truncated,
                columnCount: columns.length,
                note: truncated
                  ? `Staged the first ${rows.length} rows into ${targetRange}; the query had more. Awaiting user review and approval.`
                  : `Staged ${rows.length} rows into ${targetRange}. Awaiting user review and approval.`,
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

  private queryTimeoutMs(): number {
    const configured = this.config.options?.queryTimeoutMs;
    return typeof configured === 'number' && Number.isFinite(configured) && configured > 0
      ? configured
      : DEFAULT_QUERY_TIMEOUT_MS;
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
