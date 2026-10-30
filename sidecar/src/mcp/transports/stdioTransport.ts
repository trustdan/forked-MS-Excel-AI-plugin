import { spawn, type ChildProcess } from 'node:child_process';
import { createInterface } from 'node:readline';
import { join } from 'node:path';
import type {
  McpServerConfig,
  McpServerStatus,
  McpToolDefinition,
  McpToolResult,
} from '@excelai/shared';
import type { McpClient } from '../types.js';
import { planSpawn } from './spawnCommand.js';

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

export interface StdioLimits {
  /** Deadline for spawn + `initialize` + `tools/list` together. */
  handshakeTimeoutMs: number;
  /** Deadline for any later request, such as `tools/call`. */
  requestTimeoutMs: number;
  /** Requests awaiting a response; more are refused rather than queued. */
  maxPending: number;
  /** How long close() waits for the child to exit after asking it to stop. */
  exitGraceMs: number;
}

export const DEFAULT_STDIO_LIMITS: StdioLimits = {
  handshakeTimeoutMs: 30_000,
  requestTimeoutMs: 120_000,
  maxPending: 32,
  exitGraceMs: 2_000,
};

/** A deadline expired; raised as-is, unlike a caller's own abort. */
export class McpTimeoutError extends Error {}

interface Pending {
  resolve: (value: unknown) => void;
  reject: (err: Error) => void;
}

export class StdioMcpClient implements McpClient {
  readonly id: string;
  readonly name: string;
  readonly config: McpServerConfig;
  private readonly limits: StdioLimits;
  private process: ChildProcess | null = null;
  private nextId = 1;
  private pendingRequests = new Map<number | string, Pending>();
  private statusState: 'connected' | 'disconnected' | 'error' = 'disconnected';
  private lastError?: string;
  private cachedTools: McpToolDefinition[] = [];
  private initializing: Promise<void> | null = null;

  constructor(config: McpServerConfig, limits: Partial<StdioLimits> = {}) {
    this.id = config.id;
    this.name = config.name;
    this.config = config;
    this.limits = { ...DEFAULT_STDIO_LIMITS, ...limits };
  }

  initialize(): Promise<void> {
    if (this.process && this.statusState === 'connected') return Promise.resolve();
    // Concurrent callers share one spawn instead of racing to start two children.
    this.initializing ??= this.start().finally(() => {
      this.initializing = null;
    });
    return this.initializing;
  }

  private async start(): Promise<void> {
    const command = this.config.command;
    if (!command) {
      throw new Error(`Command is required for stdio MCP server: ${this.name}`);
    }

    // A child left over from a failed or crashed session is never reused.
    if (this.process) await this.close();

    const handshake = new AbortController();
    const timer = setTimeout(
      () => handshake.abort(new McpTimeoutError(`MCP server "${this.name}" did not complete its handshake within ${this.limits.handshakeTimeoutMs} ms`)),
      this.limits.handshakeTimeoutMs,
    );
    try {
      const env = {
        ...process.env,
        ...(this.config.env ?? {}),
      };
      const plan = planSpawn(command, this.config.args ?? [], env);
      const proc = spawn(plan.file, plan.args, {
        env,
        stdio: ['pipe', 'pipe', 'pipe'],
        windowsHide: true,
        windowsVerbatimArguments: plan.windowsVerbatimArguments,
      });
      this.process = proc;
      this.attach(proc);

      // 1. Perform initialize handshake
      await this.request(
        'initialize',
        {
          protocolVersion: '2024-11-05',
          capabilities: {},
          clientInfo: {
            name: 'excel-ai-sidecar',
            version: '0.1.0',
          },
        },
        { signal: handshake.signal, timeoutMs: 0 },
      );

      // 2. Send initialized notification
      this.notify('notifications/initialized');

      // 3. Cache available tools
      const toolsRes = (await this.request('tools/list', {}, { signal: handshake.signal, timeoutMs: 0 })) as {
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
      await this.close();
      this.statusState = 'error';
      this.lastError = e instanceof Error ? e.message : String(e);
      throw e;
    } finally {
      clearTimeout(timer);
    }
  }

  /** Wire one child's events; each handler ignores a child that is no longer current. */
  private attach(proc: ChildProcess): void {
    proc.on('error', (err) => {
      if (this.process !== proc) return;
      this.statusState = 'error';
      this.lastError = err.message;
      this.rejectAllPending(err);
    });

    proc.on('exit', (code, signal) => {
      if (this.process !== proc) return;
      this.process = null;
      this.statusState = code === 0 ? 'disconnected' : 'error';
      this.lastError = code === 0 ? undefined : `Process exited with code ${code ?? 'none'} (${signal ?? 'none'})`;
      this.rejectAllPending(new Error(`MCP server "${this.name}" exited (${code ?? signal})`));
    });

    // A write to a child that just died surfaces as an async EPIPE on stdin.
    // Unhandled, that would crash the sidecar; the exit handler reports it.
    proc.stdin?.on('error', () => {});

    if (proc.stdout) {
      const rl = createInterface({ input: proc.stdout });
      rl.on('line', (line) => this.handleStdoutLine(line));
    }
    // Drain stderr so a chatty child never blocks on a full pipe.
    proc.stderr?.resume();
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
      { signal: options?.abortSignal },
    )) as {
      content?: Array<{ type: 'text'; text: string }>;
      isError?: boolean;
    };

    return {
      content: res.content ?? [{ type: 'text', text: JSON.stringify(res) }],
      isError: res.isError ?? false,
    };
  }

  /**
   * Send a JSON-RPC request. It settles exactly once: on the response, on
   * abort, on its deadline (`timeoutMs`, 0 for none), or when the child exits.
   */
  private request(
    method: string,
    params?: Record<string, unknown>,
    options: { signal?: AbortSignal; timeoutMs?: number } = {},
  ): Promise<unknown> {
    const proc = this.process;
    if (!proc || !proc.stdin || proc.stdin.destroyed) {
      return Promise.reject(new Error(`MCP server "${this.name}" process is not running.`));
    }
    if (this.pendingRequests.size >= this.limits.maxPending) {
      return Promise.reject(
        new Error(`MCP server "${this.name}" has ${this.limits.maxPending} requests in flight; try again later.`),
      );
    }
    const { signal } = options;
    if (signal?.aborted) {
      return Promise.reject(new Error('Operation aborted'));
    }

    const id = this.nextId++;
    const payload: JsonRpcRequest = {
      jsonrpc: '2.0',
      id,
      method,
      params,
    };

    return new Promise((resolve, reject) => {
      const timeoutMs = options.timeoutMs ?? this.limits.requestTimeoutMs;
      let timer: NodeJS.Timeout | undefined;

      const settle = (fn: () => void): void => {
        if (!this.pendingRequests.delete(id)) return;
        if (timer) clearTimeout(timer);
        signal?.removeEventListener('abort', onAbort);
        fn();
      };
      // Tell the server to stop work nobody is waiting for any more.
      const giveUp = (err: Error): void =>
        settle(() => {
          this.notify('notifications/cancelled', { requestId: id, reason: err.message });
          reject(err);
        });
      const onAbort = (): void => {
        const reason = signal?.reason;
        giveUp(reason instanceof McpTimeoutError ? reason : new Error('Operation aborted'));
      };

      this.pendingRequests.set(id, {
        resolve: (val) => settle(() => resolve(val)),
        reject: (err) => settle(() => reject(err)),
      });
      signal?.addEventListener('abort', onAbort, { once: true });
      if (timeoutMs > 0) {
        timer = setTimeout(
          () => giveUp(new McpTimeoutError(`MCP server "${this.name}" did not answer "${method}" within ${timeoutMs} ms`)),
          timeoutMs,
        );
      }

      try {
        proc.stdin!.write(JSON.stringify(payload) + '\n');
      } catch (err) {
        settle(() => reject(err instanceof Error ? err : new Error(String(err))));
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
      const pending = parsed.id !== undefined ? this.pendingRequests.get(parsed.id) : undefined;
      if (pending) {
        if (parsed.error) {
          pending.reject(new Error(`MCP Error [${parsed.error.code}]: ${parsed.error.message}`));
        } else {
          pending.resolve(parsed.result);
        }
      }
    } catch {
      // not a valid JSON-RPC message, ignore
    }
  }

  private rejectAllPending(err: Error): void {
    for (const { reject } of [...this.pendingRequests.values()]) {
      reject(err);
    }
  }

  /** Stop the child and wait until it (and on Windows, its process tree) is gone. */
  async close(): Promise<void> {
    this.rejectAllPending(new Error('MCP Client closed'));
    const proc = this.process;
    this.process = null;
    this.statusState = 'disconnected';
    // No pid means the spawn itself failed, so there is no exit to wait for.
    if (!proc || proc.pid === undefined || proc.exitCode !== null || proc.signalCode !== null) return;

    const exited = new Promise<void>((resolve) => proc.once('exit', () => resolve()));
    const waitExit = (ms: number): Promise<boolean> =>
      Promise.race([
        exited.then(() => true),
        new Promise<boolean>((r) => setTimeout(() => r(false), ms).unref()),
      ]);

    proc.stdin?.end();
    if (process.platform === 'win32') {
      // TerminateProcess on the direct child would orphan anything it started
      // (cmd.exe → npx → node), so kill the whole tree.
      await killTree(proc.pid);
      await waitExit(this.limits.exitGraceMs);
      return;
    }
    proc.kill('SIGTERM');
    if (!(await waitExit(this.limits.exitGraceMs))) {
      proc.kill('SIGKILL');
      await waitExit(this.limits.exitGraceMs);
    }
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

function killTree(pid: number | undefined): Promise<void> {
  if (pid === undefined) return Promise.resolve();
  const taskkill = join(process.env.SystemRoot ?? 'C:\\Windows', 'System32', 'taskkill.exe');
  return new Promise((resolve) => {
    const killer = spawn(taskkill, ['/pid', String(pid), '/T', '/F'], { stdio: 'ignore', windowsHide: true });
    killer.once('error', () => resolve());
    killer.once('exit', () => resolve());
  });
}
