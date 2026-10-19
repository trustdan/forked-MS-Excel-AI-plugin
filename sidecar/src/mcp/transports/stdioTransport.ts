import { spawn, type ChildProcess } from 'node:child_process';
import { createInterface } from 'node:readline';
import type {
  McpServerConfig,
  McpServerStatus,
  McpToolDefinition,
  McpToolResult,
} from '@excelai/shared';
import type { McpClient } from '../types.js';

interface JsonRpcRequest {
  jsonrpc: '2.0';
  id?: number | string;
  method: string;
  params?: Record<string, unknown>;
}

interface JsonRpcResponse {
  jsonrpc: '2.0';
  id?: number | string;
  result?: unknown;
  error?: { code: number; message: string; data?: unknown };
}

export class StdioMcpClient implements McpClient {
  readonly id: string;
  readonly name: string;
  readonly config: McpServerConfig;
  private process: ChildProcess | null = null;
  private nextId = 1;
  private pendingRequests = new Map<
    number | string,
    { resolve: (value: unknown) => void; reject: (err: Error) => void }
  >();
  private statusState: 'connected' | 'disconnected' | 'error' = 'disconnected';
  private lastError?: string;
  private cachedTools: McpToolDefinition[] = [];

  constructor(config: McpServerConfig) {
    this.id = config.id;
    this.name = config.name;
    this.config = config;
  }

  async initialize(): Promise<void> {
    if (this.process) return;
    const command = this.config.command;
    if (!command) {
      throw new Error(`Command is required for stdio MCP server: ${this.name}`);
    }

    try {
      const args = this.config.args ?? [];
      const env = {
        ...process.env,
        ...(this.config.env ?? {}),
      };

      this.process = spawn(command, args, {
        env,
        stdio: ['pipe', 'pipe', 'pipe'],
        shell: process.platform === 'win32',
      });

      this.process.on('error', (err) => {
        this.statusState = 'error';
        this.lastError = err.message;
        this.rejectAllPending(err);
      });

      this.process.on('exit', (code, signal) => {
        this.statusState = 'disconnected';
        if (code !== 0 && code !== null) {
          this.lastError = `Process exited with code ${code} (${signal ?? 'none'})`;
        }
        this.rejectAllPending(new Error(`MCP process exited (${code})`));
        this.process = null;
      });

      if (this.process.stdout) {
        const rl = createInterface({ input: this.process.stdout });
        rl.on('line', (line) => this.handleStdoutLine(line));
      }

      if (this.process.stderr) {
        this.process.stderr.on('data', (data: Buffer) => {
          // Keep debug logging for stderr
          const msg = data.toString('utf8').trim();
          if (msg) {
            // Can be useful for diagnostics
          }
        });
      }

      // 1. Perform initialize handshake
      const initResult = (await this.request('initialize', {
        protocolVersion: '2024-11-05',
        capabilities: {},
        clientInfo: {
          name: 'excel-ai-sidecar',
          version: '0.1.0',
        },
      })) as Record<string, unknown>;

      // 2. Send initialized notification
      this.notify('notifications/initialized');

      // 3. Cache available tools
      const toolsRes = (await this.request('tools/list', {})) as {
        tools?: Array<{ name: string; description?: string; inputSchema?: Record<string, unknown> }>;
      };

      this.cachedTools = (toolsRes.tools ?? []).map((t) => ({
        name: t.name,
        description: t.description,
        inputSchema: t.inputSchema ?? { type: 'object', properties: {} },
      }));

      this.statusState = 'connected';
      this.lastError = undefined;
    } catch (e) {
      this.statusState = 'error';
      this.lastError = e instanceof Error ? e.message : String(e);
      await this.close();
      throw e;
    }
  }

  async listTools(): Promise<McpToolDefinition[]> {
    if (!this.process) {
      await this.initialize();
    }
    return this.cachedTools;
  }

  async callTool(
    toolName: string,
    args: Record<string, unknown>,
    options?: { abortSignal?: AbortSignal },
  ): Promise<McpToolResult> {
    if (!this.process) {
      await this.initialize();
    }

    if (options?.abortSignal?.aborted) {
      throw new Error('Call aborted');
    }

    const res = (await this.request(
      'tools/call',
      {
        name: toolName,
        arguments: args,
      },
      options?.abortSignal,
    )) as {
      content?: Array<{ type: 'text'; text: string }>;
      isError?: boolean;
    };

    return {
      content: res.content ?? [{ type: 'text', text: JSON.stringify(res) }],
      isError: res.isError ?? false,
    };
  }

  private request(
    method: string,
    params?: Record<string, unknown>,
    abortSignal?: AbortSignal,
  ): Promise<unknown> {
    if (!this.process || !this.process.stdin || this.process.stdin.destroyed) {
      return Promise.reject(new Error(`MCP server "${this.name}" process is not running.`));
    }

    const id = this.nextId++;
    const payload: JsonRpcRequest = {
      jsonrpc: '2.0',
      id,
      method,
      params,
    };

    return new Promise((resolve, reject) => {
      let abortHandler: (() => void) | undefined;
      if (abortSignal) {
        if (abortSignal.aborted) {
          return reject(new Error('Operation aborted'));
        }
        abortHandler = () => {
          this.pendingRequests.delete(id);
          reject(new Error('Operation aborted'));
        };
        abortSignal.addEventListener('abort', abortHandler, { once: true });
      }

      this.pendingRequests.set(id, {
        resolve: (val) => {
          if (abortHandler && abortSignal) {
            abortSignal.removeEventListener('abort', abortHandler);
          }
          resolve(val);
        },
        reject: (err) => {
          if (abortHandler && abortSignal) {
            abortSignal.removeEventListener('abort', abortHandler);
          }
          reject(err);
        },
      });

      try {
        this.process!.stdin!.write(JSON.stringify(payload) + '\n');
      } catch (err) {
        this.pendingRequests.delete(id);
        reject(err);
      }
    });
  }

  private notify(method: string, params?: Record<string, unknown>): void {
    if (!this.process || !this.process.stdin || this.process.stdin.destroyed) return;
    const payload: JsonRpcRequest = {
      jsonrpc: '2.0',
      method,
      params,
    };
    try {
      this.process.stdin.write(JSON.stringify(payload) + '\n');
    } catch {
      // ignore write error on notification
    }
  }

  private handleStdoutLine(line: string): void {
    const trimmed = line.trim();
    if (!trimmed.startsWith('{')) return;
    try {
      const parsed = JSON.parse(trimmed) as JsonRpcResponse;
      if (parsed.id !== undefined && this.pendingRequests.has(parsed.id)) {
        const { resolve, reject } = this.pendingRequests.get(parsed.id)!;
        this.pendingRequests.delete(parsed.id);
        if (parsed.error) {
          reject(new Error(`MCP Error [${parsed.error.code}]: ${parsed.error.message}`));
        } else {
          resolve(parsed.result);
        }
      }
    } catch {
      // not a valid JSON-RPC message, ignore
    }
  }

  private rejectAllPending(err: Error): void {
    for (const { reject } of this.pendingRequests.values()) {
      reject(err);
    }
    this.pendingRequests.clear();
  }

  async close(): Promise<void> {
    this.rejectAllPending(new Error('MCP Client closed'));
    if (this.process) {
      const proc = this.process;
      this.process = null;
      try {
        proc.kill('SIGTERM');
        // Give it 500ms before SIGKILL
        setTimeout(() => {
          if (!proc.killed) {
            try {
              proc.kill('SIGKILL');
            } catch {}
          }
        }, 500).unref();
      } catch {}
    }
    this.statusState = 'disconnected';
  }

  getStatus(): McpServerStatus {
    return {
      id: this.id,
      name: this.name,
      enabled: this.config.enabled,
      transport: 'stdio',
      status: this.statusState,
      error: this.lastError,
      toolCount: this.cachedTools.length,
      tools: this.cachedTools.map((t) => ({ name: t.name, description: t.description })),
    };
  }
}
