import { readdirSync, statSync, existsSync } from 'node:fs';
import { join, extname, relative } from 'node:path';
import type {
  McpServerConfig,
  McpServerStatus,
  McpToolDefinition,
  McpToolResult,
} from '@excelai/shared';
import { PathGuard } from '../pathGuard.js';
import type { McpClient } from '../types.js';
import {
  generateImportChangeOps,
  loadAndParseLocalFile,
  parseTabularContent,
} from '../../importers/tabular.js';

export class CsvDirectoryConnector implements McpClient {
  readonly id: string;
  readonly name: string;
  readonly config: McpServerConfig;
  private dirPath: string | null = null;
  private statusState: 'connected' | 'disconnected' | 'error' = 'disconnected';
  private lastError?: string;

  constructor(config: McpServerConfig) {
    this.id = config.id;
    this.name = config.name;
    this.config = config;
  }

  async initialize(): Promise<void> {
    try {
      const rawDirPath = String(this.config.options?.dirPath ?? '');
      if (!rawDirPath) {
        throw new Error('options.dirPath is required for CSV Directory connector');
      }
      this.dirPath = PathGuard.assertApproved(
        rawDirPath,
        this.config.allowedPaths,
        `CSV Directory Connector [${this.name}]`,
      );

      if (!existsSync(this.dirPath)) {
        throw new Error(`Directory does not exist: ${this.dirPath}`);
      }
      const stat = statSync(this.dirPath);
      if (!stat.isDirectory()) {
        throw new Error(`Path is not a directory: ${this.dirPath}`);
      }

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
        name: 'csv_list_files',
        description: 'List all CSV and TSV files in the approved directory, including file sizes and modification dates.',
        inputSchema: {
          type: 'object',
          properties: {
            subDir: {
              type: 'string',
              description: 'Optional subfolder within the CSV directory to search',
            },
          },
          additionalProperties: false,
        },
      },
      {
        name: 'csv_read_preview',
        description: 'Read the header and top N rows of a CSV or TSV file to inspect its schema and sample values.',
        inputSchema: {
          type: 'object',
          properties: {
            relativePath: {
              type: 'string',
              description: 'Relative path of the CSV or TSV file within the approved directory',
            },
            maxRows: {
              type: 'number',
              description: 'Maximum preview rows to return (default 10, max 100)',
            },
          },
          required: ['relativePath'],
          additionalProperties: false,
        },
      },
      {
        name: 'csv_query',
        description: 'Search and filter rows in a CSV or TSV file by column name and value substring.',
        inputSchema: {
          type: 'object',
          properties: {
            relativePath: {
              type: 'string',
              description: 'Relative path of the CSV or TSV file',
            },
            column: {
              type: 'string',
              description: 'Column name to filter by (or empty to search all columns)',
            },
            filterValue: {
              type: 'string',
              description: 'Value or substring to match',
            },
            limit: {
              type: 'number',
              description: 'Maximum matching rows to return (default 100, max 500)',
            },
          },
          required: ['relativePath'],
          additionalProperties: false,
        },
      },
      {
        name: 'csv_stage_to_sheet',
        description: 'Parse a CSV or TSV file from the approved directory and stage its full content into an Excel worksheet change-set for user review.',
        inputSchema: {
          type: 'object',
          properties: {
            relativePath: {
              type: 'string',
              description: 'Relative path of the CSV or TSV file',
            },
            targetSheet: {
              type: 'string',
              description: 'Target worksheet name (defaults to "Sheet1")',
            },
            startCell: {
              type: 'string',
              description: 'Top-left cell address (defaults to "A1")',
            },
            hasHeaders: {
              type: 'boolean',
              description: 'Whether row 1 contains headers (defaults to true)',
            },
            reason: {
              type: 'string',
              description: 'One-line justification for staging this CSV import',
            },
          },
          required: ['relativePath'],
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
    if (!this.dirPath) {
      await this.initialize();
    }
    const root = this.dirPath!;

    switch (toolName) {
      case 'csv_list_files': {
        const sub = String(args.subDir ?? '').trim();
        const targetDir = PathGuard.assertApproved(
          sub ? join(root, sub) : root,
          this.config.allowedPaths,
          'csv_list_files',
        );

        const files: Array<{ relativePath: string; sizeBytes: number; modifiedAt: string }> = [];
        const scan = (currentDir: string): void => {
          const entries = readdirSync(currentDir, { withFileTypes: true });
          for (const ent of entries) {
            const full = join(currentDir, ent.name);
            if (ent.isDirectory()) {
              // Recurse if safe
              if (PathGuard.isApproved(full, this.config.allowedPaths)) {
                scan(full);
              }
            } else if (ent.isFile()) {
              const ext = extname(ent.name).toLowerCase();
              if (ext === '.csv' || ext === '.tsv') {
                const stat = statSync(full);
                files.push({
                  relativePath: relative(root, full).replace(/\\/g, '/'),
                  sizeBytes: stat.size,
                  modifiedAt: stat.mtime.toISOString(),
                });
              }
            }
          }
        };

        scan(targetDir);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ ok: true, totalFiles: files.length, files }),
            },
          ],
        };
      }

      case 'csv_read_preview': {
        const relPath = String(args.relativePath ?? '').trim();
        if (!relPath) {
          return { content: [{ type: 'text', text: 'Error: relativePath is required' }], isError: true };
        }
        const fullPath = PathGuard.assertApproved(
          join(root, relPath),
          this.config.allowedPaths,
          'csv_read_preview',
        );
        const maxRows = Math.min(Math.max(1, Number(args.maxRows ?? 10)), 100);

        const parsed = loadAndParseLocalFile(fullPath);
        const previewRows = parsed.rows.slice(0, maxRows);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                ok: true,
                filePath: relPath,
                format: parsed.format,
                totalRows: parsed.totalRows,
                totalCols: parsed.totalCols,
                previewRowCount: previewRows.length,
                rows: previewRows,
              }),
            },
          ],
        };
      }

      case 'csv_query': {
        const relPath = String(args.relativePath ?? '').trim();
        if (!relPath) {
          return { content: [{ type: 'text', text: 'Error: relativePath is required' }], isError: true };
        }
        const fullPath = PathGuard.assertApproved(
          join(root, relPath),
          this.config.allowedPaths,
          'csv_query',
        );
        const colFilter = String(args.column ?? '').trim().toLowerCase();
        const filterVal = String(args.filterValue ?? '').trim().toLowerCase();
        const limit = Math.min(Math.max(1, Number(args.limit ?? 100)), 500);

        const parsed = loadAndParseLocalFile(fullPath);
        if (parsed.rows.length === 0) {
          return {
            content: [{ type: 'text', text: JSON.stringify({ ok: true, matchingRows: 0, rows: [] }) }],
          };
        }

        const headers = (parsed.rows[0] ?? []).map(String);
        let colIndex = -1;
        if (colFilter) {
          colIndex = headers.findIndex((h) => h.toLowerCase() === colFilter);
          if (colIndex === -1) {
            return {
              content: [
                {
                  type: 'text',
                  text: JSON.stringify({
                    ok: false,
                    error: `Column "${colFilter}" not found in headers: [${headers.join(', ')}]`,
                  }),
                },
              ],
              isError: true,
            };
          }
        }

        const matched: (string | number | boolean | null)[][] = [headers];
        const dataRows = parsed.rows.slice(1);
        for (const row of dataRows) {
          if (matched.length > limit) break;
          if (!filterVal) {
            matched.push(row);
            continue;
          }
          if (colIndex >= 0) {
            const cellVal = String(row[colIndex] ?? '').toLowerCase();
            if (cellVal.includes(filterVal)) {
              matched.push(row);
            }
          } else {
            // Search all columns
            const anyMatch = row.some((c) => String(c ?? '').toLowerCase().includes(filterVal));
            if (anyMatch) {
              matched.push(row);
            }
          }
        }

        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                ok: true,
                totalMatched: matched.length - 1,
                columns: headers,
                rows: matched,
              }),
            },
          ],
        };
      }

      case 'csv_stage_to_sheet': {
        const relPath = String(args.relativePath ?? '').trim();
        if (!relPath) {
          return { content: [{ type: 'text', text: 'Error: relativePath is required' }], isError: true };
        }
        const fullPath = PathGuard.assertApproved(
          join(root, relPath),
          this.config.allowedPaths,
          'csv_stage_to_sheet',
        );
        const targetSheet = String(args.targetSheet ?? 'Sheet1').trim();
        const startCell = String(args.startCell ?? 'A1').trim();
        const hasHeaders = args.hasHeaders !== false;
        const reason = String(args.reason ?? `Import CSV: ${relPath}`);

        const parsed = loadAndParseLocalFile(fullPath);
        const ops = generateImportChangeOps({
          rows: parsed.rows,
          targetSheet,
          startCell,
          hasHeaders,
          reason,
        });

        const targetRange = ops[0]?.kind === 'write_range' ? ops[0].range : undefined;
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                ok: true,
                staged: true,
                format: parsed.format,
                rowsCount: parsed.totalRows,
                colsCount: parsed.totalCols,
                targetRange,
                note: `Staged ${parsed.totalRows} rows × ${parsed.totalCols} cols from "${relPath}" into pending change-set. Awaiting approval.`,
              }),
            },
          ],
          changeOps: ops,
        };
      }

      default:
        return {
          content: [{ type: 'text', text: `Unknown tool: ${toolName}` }],
          isError: true,
        };
    }
  }

  async close(): Promise<void> {
    this.statusState = 'disconnected';
  }

  getStatus(): McpServerStatus {
    return {
      id: this.id,
      name: this.name,
      enabled: this.config.enabled,
      transport: 'builtin',
      type: 'csv-directory',
      status: this.statusState,
      error: this.lastError,
      toolCount: 4,
    };
  }
}
