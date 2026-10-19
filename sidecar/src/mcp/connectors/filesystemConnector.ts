import { readdirSync, statSync, readFileSync, existsSync } from 'node:fs';
import { join, relative } from 'node:path';
import type {
  McpServerConfig,
  McpServerStatus,
  McpToolDefinition,
  McpToolResult,
} from '@excelai/shared';
import { PathGuard } from '../pathGuard.js';
import type { McpClient } from '../types.js';

const MAX_FILE_READ_BYTES = 256 * 1024; // 256 KB max for text reads

export class FilesystemConnector implements McpClient {
  readonly id: string;
  readonly name: string;
  readonly config: McpServerConfig;
  private rootPath: string | null = null;
  private statusState: 'connected' | 'disconnected' | 'error' = 'disconnected';
  private lastError?: string;

  constructor(config: McpServerConfig) {
    this.id = config.id;
    this.name = config.name;
    this.config = config;
  }

  async initialize(): Promise<void> {
    try {
      const rawRoot = String(this.config.options?.rootPath ?? '');
      if (!rawRoot) {
        throw new Error('options.rootPath is required for Filesystem connector');
      }
      this.rootPath = PathGuard.assertApproved(
        rawRoot,
        this.config.allowedPaths,
        `Filesystem Connector [${this.name}]`,
      );

      if (!existsSync(this.rootPath)) {
        throw new Error(`Root path does not exist: ${this.rootPath}`);
      }
      const stat = statSync(this.rootPath);
      if (!stat.isDirectory()) {
        throw new Error(`Root path is not a directory: ${this.rootPath}`);
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
        name: 'fs_list_directory',
        description: 'List contents of a directory within the approved filesystem root.',
        inputSchema: {
          type: 'object',
          properties: {
            path: {
              type: 'string',
              description: 'Relative path within the approved root (empty for root)',
            },
          },
          additionalProperties: false,
        },
      },
      {
        name: 'fs_read_file',
        description: 'Read the text content of a file within the approved filesystem root (up to 256 KB).',
        inputSchema: {
          type: 'object',
          properties: {
            path: {
              type: 'string',
              description: 'Relative path of the file to read',
            },
            maxBytes: {
              type: 'number',
              description: 'Maximum bytes to read (default 65536, max 262144)',
            },
          },
          required: ['path'],
          additionalProperties: false,
        },
      },
      {
        name: 'fs_search_files',
        description: 'Search for files by name or extension within the approved directory tree.',
        inputSchema: {
          type: 'object',
          properties: {
            query: {
              type: 'string',
              description: 'Search term to match against file names',
            },
            maxResults: {
              type: 'number',
              description: 'Maximum matches to return (default 50, max 100)',
            },
          },
          required: ['query'],
          additionalProperties: false,
        },
      },
      {
        name: 'fs_get_metadata',
        description: 'Get file system metadata for a path (size, timestamps, type).',
        inputSchema: {
          type: 'object',
          properties: {
            path: {
              type: 'string',
              description: 'Relative path within the approved root',
            },
          },
          required: ['path'],
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
    if (!this.rootPath) {
      await this.initialize();
    }
    const root = this.rootPath!;

    switch (toolName) {
      case 'fs_list_directory': {
        const sub = String(args.path ?? '').trim();
        const targetDir = PathGuard.assertApproved(
          sub ? join(root, sub) : root,
          this.config.allowedPaths,
          'fs_list_directory',
        );

        if (!existsSync(targetDir)) {
          return { content: [{ type: 'text', text: `Directory not found: ${sub}` }], isError: true };
        }

        const entries = readdirSync(targetDir, { withFileTypes: true });
        const items = entries.map((ent) => {
          const full = join(targetDir, ent.name);
          let size = 0;
          try {
            if (ent.isFile()) {
              size = statSync(full).size;
            }
          } catch {
            // ignore inaccessible
          }
          return {
            name: ent.name,
            isDirectory: ent.isDirectory(),
            isFile: ent.isFile(),
            sizeBytes: size,
            relativePath: relative(root, full).replace(/\\/g, '/'),
          };
        });

        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ ok: true, path: sub || '.', count: items.length, items }),
            },
          ],
        };
      }

      case 'fs_read_file': {
        const rel = String(args.path ?? '').trim();
        if (!rel) {
          return { content: [{ type: 'text', text: 'Error: path is required' }], isError: true };
        }
        const fullPath = PathGuard.assertApproved(
          join(root, rel),
          this.config.allowedPaths,
          'fs_read_file',
        );

        if (!existsSync(fullPath)) {
          return { content: [{ type: 'text', text: `File not found: ${rel}` }], isError: true };
        }
        const stat = statSync(fullPath);
        if (!stat.isFile()) {
          return { content: [{ type: 'text', text: `Path is not a regular file: ${rel}` }], isError: true };
        }

        const maxBytes = Math.min(Math.max(1, Number(args.maxBytes ?? 64 * 1024)), MAX_FILE_READ_BYTES);
        const buf = readFileSync(fullPath);
        const isTruncated = buf.length > maxBytes;
        const slice = isTruncated ? buf.subarray(0, maxBytes) : buf;

        // Check for binary null byte
        if (slice.includes(0)) {
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  ok: false,
                  error: 'Cannot read binary file as text.',
                  fileSizeBytes: stat.size,
                }),
              },
            ],
            isError: true,
          };
        }

        const text = slice.toString('utf8');
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                ok: true,
                path: rel,
                sizeBytes: stat.size,
                bytesRead: slice.length,
                truncated: isTruncated,
                content: text,
              }),
            },
          ],
        };
      }

      case 'fs_search_files': {
        const q = String(args.query ?? '').trim().toLowerCase();
        if (!q) {
          return { content: [{ type: 'text', text: 'Error: query is required' }], isError: true };
        }
        const maxResults = Math.min(Math.max(1, Number(args.maxResults ?? 50)), 100);

        const results: Array<{ relativePath: string; isDirectory: boolean; sizeBytes: number }> = [];
        const scan = (currentDir: string): void => {
          if (results.length >= maxResults) return;
          const entries = readdirSync(currentDir, { withFileTypes: true });
          for (const ent of entries) {
            if (results.length >= maxResults) break;
            const full = join(currentDir, ent.name);
            const matches = ent.name.toLowerCase().includes(q);
            if (matches) {
              let size = 0;
              try {
                if (ent.isFile()) size = statSync(full).size;
              } catch {}
              results.push({
                relativePath: relative(root, full).replace(/\\/g, '/'),
                isDirectory: ent.isDirectory(),
                sizeBytes: size,
              });
            }
            if (ent.isDirectory() && PathGuard.isApproved(full, this.config.allowedPaths)) {
              scan(full);
            }
          }
        };

        scan(root);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ ok: true, query: q, totalMatches: results.length, matches: results }),
            },
          ],
        };
      }

      case 'fs_get_metadata': {
        const rel = String(args.path ?? '').trim();
        const fullPath = PathGuard.assertApproved(
          rel ? join(root, rel) : root,
          this.config.allowedPaths,
          'fs_get_metadata',
        );

        if (!existsSync(fullPath)) {
          return { content: [{ type: 'text', text: `Path not found: ${rel}` }], isError: true };
        }
        const stat = statSync(fullPath);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                ok: true,
                path: rel || '.',
                sizeBytes: stat.size,
                isDirectory: stat.isDirectory(),
                isFile: stat.isFile(),
                createdAt: stat.birthtime.toISOString(),
                modifiedAt: stat.mtime.toISOString(),
              }),
            },
          ],
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
      type: 'filesystem',
      status: this.statusState,
      error: this.lastError,
      toolCount: 4,
    };
  }
}
