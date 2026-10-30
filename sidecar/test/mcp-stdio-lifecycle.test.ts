import { mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { dataDir, saveSettings } from '../src/config.js';
import { openDb } from '../src/db.js';
import { McpManager } from '../src/mcp/manager.js';
import { planSpawn } from '../src/mcp/transports/spawnCommand.js';
import { StdioMcpClient, type StdioLimits } from '../src/mcp/transports/stdioTransport.js';
import { startServer, type SidecarServer } from '../src/server.js';
import type { McpServerConfig } from '@excelai/shared';

/**
 * Mock MCP server. MOCK_MODE picks its behaviour:
 *   normal     answers everything
 *   silent     never answers (a child that hangs at startup)
 *   hang-call  completes the handshake, never answers tools/call
 *   exit-call  exits with code 3 on tools/call
 * Every received method, and its own pid, is appended to MOCK_LOG.
 */
const SERVER_SOURCE = `
const fs = require('fs');
const readline = require('readline');
const mode = process.env.MOCK_MODE || 'normal';
const log = (s) => fs.appendFileSync(process.env.MOCK_LOG, s + '\\n');
log('pid:' + process.pid);
const send = (id, result) => process.stdout.write(JSON.stringify({ jsonrpc: '2.0', id, result }) + '\\n');
readline.createInterface({ input: process.stdin }).on('line', (line) => {
  const req = JSON.parse(line);
  log(req.method);
  if (mode === 'silent') return;
  if (req.method === 'initialize') send(req.id, { protocolVersion: '2024-11-05', capabilities: {} });
  else if (req.method === 'tools/list') send(req.id, { tools: [{ name: 'argv' }] });
  else if (req.method === 'tools/call') {
    if (mode === 'hang-call') return;
    if (mode === 'exit-call') process.exit(3);
    send(req.id, { content: [{ type: 'text', text: JSON.stringify({ argv: process.argv.slice(2), pid: process.pid }) }] });
  }
});
`;

let dir: string;
let script: string;
let logFile: string;
const clients: StdioMcpClient[] = [];

beforeEach(() => {
  dir = join(tmpdir(), `excelai mcp ${crypto.randomUUID()}`);
  mkdirSync(dir, { recursive: true });
  script = join(dir, 'server.cjs');
  logFile = join(dir, 'log.txt');
  writeFileSync(script, SERVER_SOURCE);
  writeFileSync(logFile, '');
});

afterEach(async () => {
  await Promise.all(clients.splice(0).map((c) => c.close()));
  rmSync(dir, { recursive: true, force: true });
});

function config(mode: string, overrides: Partial<McpServerConfig> = {}): McpServerConfig {
  return {
    id: `stdio-${mode}`,
    name: `Mock ${mode}`,
    enabled: true,
    transport: 'stdio',
    command: process.execPath,
    args: [script],
    env: { MOCK_MODE: mode, MOCK_LOG: logFile },
    allowedPaths: [],
    ...overrides,
  };
}

function client(mode: string, limits: Partial<StdioLimits> = {}, overrides: Partial<McpServerConfig> = {}) {
  const c = new StdioMcpClient(config(mode, overrides), limits);
  clients.push(c);
  return c;
}

const logLines = (): string[] => readFileSync(logFile, 'utf8').split('\n').filter(Boolean);
const pids = (): number[] => logLines().filter((l) => l.startsWith('pid:')).map((l) => Number(l.slice(4)));

function isAlive(pid: number): boolean {
  try {
    process.kill(pid, 0);
    return true;
  } catch {
    return false;
  }
}

async function until(check: () => boolean, ms = 5_000): Promise<void> {
  const end = Date.now() + ms;
  while (!check()) {
    if (Date.now() > end) throw new Error('condition not met in time');
    await new Promise((r) => setTimeout(r, 20));
  }
}

describe('StdioMcpClient deadlines and caps (L2)', () => {
  it('fails the handshake of a silent child by its deadline and stops the child', async () => {
    const c = client('silent', { handshakeTimeoutMs: 300 });
    const started = Date.now();

    await expect(c.initialize()).rejects.toThrow(/did not complete its handshake within 300 ms/);
    expect(Date.now() - started).toBeLessThan(5_000);
    expect(c.getStatus()).toMatchObject({ status: 'error', error: expect.stringContaining('handshake') });
    expect(pids()).toHaveLength(1);
    expect(isAlive(pids()[0]!)).toBe(false);
  });

  it('fails a tool call by its deadline and tells the server to cancel it', async () => {
    const c = client('hang-call', { requestTimeoutMs: 300 });
    await c.initialize();

    await expect(c.callTool('argv', {})).rejects.toThrow(/did not answer "tools\/call" within 300 ms/);
    await until(() => logLines().includes('notifications/cancelled'));
  });

  it('refuses requests beyond the pending cap instead of queueing them', async () => {
    const c = client('hang-call', { maxPending: 2 });
    await c.initialize();

    const first = c.callTool('argv', {});
    const second = c.callTool('argv', {});
    first.catch(() => {});
    second.catch(() => {});
    await expect(c.callTool('argv', {})).rejects.toThrow(/2 requests in flight/);

    await c.close();
    await expect(first).rejects.toThrow(/closed/);
    await expect(second).rejects.toThrow(/closed/);
  });

  it('rejects in-flight calls when the child exits, reports it, and restarts on next use', async () => {
    const c = client('exit-call');
    await c.initialize();

    await expect(c.callTool('argv', {})).rejects.toThrow(/exited \(3\)/);
    expect(c.getStatus()).toMatchObject({ status: 'error', error: expect.stringContaining('code 3') });

    await c.listTools();
    expect(c.getStatus().status).toBe('connected');
    expect(pids()).toHaveLength(2);
  });

  it('close() resolves only after the child has exited', async () => {
    const c = client('normal');
    await c.initialize();
    const [pid] = pids();
    expect(isAlive(pid!)).toBe(true);

    await c.close();
    expect(isAlive(pid!)).toBe(false);
  });

  it('starts once when initialize() is called concurrently', async () => {
    const c = client('normal');
    await Promise.all([c.initialize(), c.initialize(), c.listTools()]);
    expect(pids()).toHaveLength(1);
  });
});

describe.skipIf(process.platform !== 'win32')('Windows command handling (L2)', () => {
  /** A .cmd shim, like npx.cmd: cmd.exe runs it and it starts node with %*. */
  function shim(): string {
    const path = join(dir, 'mock server.cmd');
    writeFileSync(path, `@"${process.execPath}" "${script}" %*\r\n`);
    return path;
  }

  it('passes arguments through a .cmd shim literally, without shell interpretation', async () => {
    const args = ['two words', 'a&b|c', '(paren)', '<angle>', 'caret^', 'bang!', 'trailing\\', ''];
    const c = client('normal', {}, { command: shim(), args: [script, ...args] });
    // The shim already names the script, so it arrives once more as argv[0].
    const res = await c.callTool('argv', {});
    const { argv } = JSON.parse(res.content[0]!.text) as { argv: string[] };
    expect(argv).toEqual([script, ...args]);
  });

  it('kills the whole process tree behind a .cmd shim on close', async () => {
    const c = client('normal', {}, { command: shim(), args: [] });
    await c.initialize();
    const [nodePid] = pids();
    expect(isAlive(nodePid!)).toBe(true);

    await c.close();
    expect(isAlive(nodePid!)).toBe(false);
  });

  it.each([['%PATH%'], ['say "hi"'], ['line\nbreak']])(
    'refuses an argument cmd.exe would reinterpret: %j',
    async (arg) => {
      const c = client('normal', {}, { command: shim(), args: [arg] });
      await expect(c.initialize()).rejects.toThrow(/cmd\.exe would reinterpret/);
      expect(pids()).toHaveLength(0);
    },
  );

  it('spawns a .exe directly and resolves bare names through PATH and PATHEXT', () => {
    const env = { Path: join(process.execPath, '..'), PATHEXT: '.COM;.EXE;.BAT;.CMD' };
    const plan = planSpawn('node', ['x y'], env, 'win32');
    expect(plan.file.toLowerCase()).toBe(process.execPath.toLowerCase());
    expect(plan.args).toEqual(['x y']);
    expect(plan.windowsVerbatimArguments).toBeUndefined();
  });
});

describe('planSpawn off Windows', () => {
  it('spawns the command as given', () => {
    expect(planSpawn('npx', ['a b', '%X%'], {}, 'linux')).toEqual({ file: 'npx', args: ['a b', '%X%'] });
  });
});

describe('MCP startup and shutdown (L2)', () => {
  let sidecar: SidecarServer | null = null;

  afterEach(async () => {
    await sidecar?.close();
    sidecar = null;
    rmSync(join(dataDir(), 'settings.json'), { force: true });
  });

  it('serves HTTP while a configured MCP child is silent, and shutdown stops that child', async () => {
    saveSettings({ providers: [], activeProviderId: null, mcpServers: [config('silent')], approvedPaths: [] });

    const started = Date.now();
    sidecar = await startServer({ db: openDb(':memory:'), port: 0 });
    expect((await fetch(`http://127.0.0.1:${sidecar.port}/api/health`)).status).toBe(200);
    expect(Date.now() - started).toBeLessThan(5_000);
    await until(() => pids().length === 1);

    const closing = Date.now();
    await sidecar.close();
    sidecar = null;
    expect(Date.now() - closing).toBeLessThan(5_000);
    expect(isAlive(pids()[0]!)).toBe(false);
  });

  it('serializes overlapping syncs so one server never gets two children', async () => {
    const manager = new McpManager();
    const cfg = config('normal');
    await Promise.all([manager.sync([cfg]), manager.sync([cfg]), manager.sync([cfg])]);
    expect(pids()).toHaveLength(1);
    await manager.closeAll();
    expect(isAlive(pids()[0]!)).toBe(false);
  });

  it('reports why a server failed to start', async () => {
    const manager = new McpManager();
    const cfg = config('normal', { command: join(dir, 'missing.exe') });
    await manager.sync([cfg]);
    const [status] = manager.getServerStatuses([cfg]);
    expect(status).toMatchObject({ status: 'error', error: expect.stringMatching(/ENOENT/) });
    await manager.closeAll();
  });
});
