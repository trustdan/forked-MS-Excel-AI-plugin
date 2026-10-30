import { readdirSync, statSync, existsSync } from 'node:fs';
import { join, relative } from 'node:path';
import { z } from 'zod';
import type {
  McpServerConfig,
  McpServerStatus,
  McpToolDefinition,
  McpToolResult,
} from '@excelai/shared';
import { PathGuard } from '../pathGuard.js';
import type { McpClient } from '../types.js';
import { parseArgs, toolDefinitions } from '../builtinTools.js';
import { readFileCapped } from '../../boundedRead.js';

const MAX_FILE_READ_BYTES = 256 * 1024; // 256 KB max for text reads

const TOOLS = {
  fs_list_directory: {
    description: 'List contents of a directory within the approved filesystem root.',
    parameters: z
      .object({
        path: z.string().optional().describe('Relative path within the approved root (empty for root)'),
      })
      .strict(),
  },
  fs_read_file: {
    description: 'Read the text content of a file within the approved filesystem root (up to 256 KB).',
    parameters: z
      .object({
        path: z.string().trim().min(1).describe('Relative path of the file to read'),
        maxBytes: z
          .number()
          .int()
          .min(1)
          .max(MAX_FILE_READ_BYTES)
          .optional()
          .describe('Maximum bytes to read (default 65536, max 262144)'),
      })
      .strict(),
  },
  fs_search_files: {
    description: 'Search for files by name or extension within the approved directory tree.',
    parameters: z
      .object({
        query: z.string().trim().min(1).describe('Search term to match against file names'),
        maxResults: z.number().int().min(1).max(100).optional().describe('Maximum matches to return (default 50, max 100)'),
      })
      .strict(),
  },
  fs_get_metadata: {
    description: 'Get file system metadata for a path (size, timestamps, type).',
    parameters: z
      .object({
        path: z.string().describe('Relative path within the approved root'),
      })
      .strict(),
  },
};

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
    return toolDefinitions(TOOLS);
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
        const sub = (parseArgs(TOOLS.fs_list_directory, args).path ?? '').trim();
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
        const { path: rel, maxBytes = 64 * 1024 } = parseArgs(TOOLS.fs_read_file, args);
        const fullPath = PathGuard.assertApproved(
          join(root, rel),
          this.config.allowedPaths,
          'fs_read_file',
        );

        if (!existsSync(fullPath)) {
          return { content: [{ type: 'text', text: `File not found: ${rel}` }], isError: true };
        }
        if (!statSync(fullPath).isFile()) {
          return { content: [{ type: 'text', text: `Path is not a regular file: ${rel}` }], isError: true };
        }

        const { bytes: slice, size, truncated: isTruncated } = readFileCapped(fullPath, maxBytes);

        // Check for binary null byte
        if (slice.includes(0)) {
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({
                  ok: false,
                  error: 'Cannot read binary file as text.',
                  fileSizeBytes: size,
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
                sizeBytes: size,
                bytesRead: slice.length,
                truncated: isTruncated,
                content: text,
              }),
            },
          ],
        };
      }

      case 'fs_search_files': {
        const { query, maxResults = 50 } = parseArgs(TOOLS.fs_search_files, args);
        const q = query.toLowerCase();

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
        const rel = parseArgs(TOOLS.fs_get_metadata, args).path.trim();
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
