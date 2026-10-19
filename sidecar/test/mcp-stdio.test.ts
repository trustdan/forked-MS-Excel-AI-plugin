import { describe, it, expect, afterEach } from 'vitest';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { writeFileSync, rmSync } from 'node:fs';
import { StdioMcpClient } from '../src/mcp/transports/stdioTransport.js';

describe('StdioMcpClient', () => {
  let serverScriptPath: string | null = null;
  let client: StdioMcpClient | null = null;

  afterEach(async () => {
    if (client) {
      await client.close();
      client = null;
    }
    if (serverScriptPath) {
      try {
        rmSync(serverScriptPath, { force: true });
      } catch {}
      serverScriptPath = null;
    }
  });

  it('performs JSON-RPC 2.0 handshake, tool discovery, and tool execution', async () => {
    serverScriptPath = join(tmpdir(), `mock-mcp-server-${Date.now()}.cjs`);
    writeFileSync(
      serverScriptPath,
      `
const readline = require('readline');
const rl = readline.createInterface({ input: process.stdin });
rl.on('line', (line) => {
  try {
    const req = JSON.parse(line);
    if (req.method === 'initialize') {
      process.stdout.write(JSON.stringify({
        jsonrpc: '2.0',
        id: req.id,
        result: {
          protocolVersion: '2024-11-05',
          capabilities: {},
          serverInfo: { name: 'mock-mcp-server', version: '1.0.0' }
        }
      }) + '\\n');
    } else if (req.method === 'notifications/initialized') {
      // acknowledged
    } else if (req.method === 'tools/list') {
      process.stdout.write(JSON.stringify({
        jsonrpc: '2.0',
        id: req.id,
        result: {
          tools: [
            {
              name: 'echo_text',
              description: 'Echoes input message',
              inputSchema: {
                type: 'object',
                properties: { message: { type: 'string' } },
                required: ['message']
              }
            }
          ]
        }
      }) + '\\n');
    } else if (req.method === 'tools/call') {
      process.stdout.write(JSON.stringify({
        jsonrpc: '2.0',
        id: req.id,
        result: {
          content: [{ type: 'text', text: 'ECHO: ' + req.params.arguments.message }]
        }
      }) + '\\n');
    }
  } catch (e) {}
});
      `,
    );

    client = new StdioMcpClient({
      id: 'stdio-test',
      name: 'Mock Stdio Server',
      enabled: true,
      transport: 'stdio',
      command: process.execPath,
      args: [serverScriptPath],
      allowedPaths: [],
    });

    await client.initialize();
    const status = client.getStatus();
    expect(status.status).toBe('connected');

    const tools = await client.listTools();
    expect(tools).toHaveLength(1);
    expect(tools[0].name).toBe('echo_text');

    const callRes = await client.callTool('echo_text', { message: 'Hello MCP!' });
    expect(callRes.isError).toBeFalsy();
    expect(callRes.content[0].text).toBe('ECHO: Hello MCP!');
  });

  it('handles cancellation via AbortSignal', async () => {
    serverScriptPath = join(tmpdir(), `mock-mcp-slow-${Date.now()}.cjs`);
    writeFileSync(
      serverScriptPath,
      `
const readline = require('readline');
const rl = readline.createInterface({ input: process.stdin });
rl.on('line', (line) => {
  const req = JSON.parse(line);
  if (req.method === 'initialize') {
    process.stdout.write(JSON.stringify({ jsonrpc: '2.0', id: req.id, result: {} }) + '\\n');
  } else if (req.method === 'tools/list') {
    process.stdout.write(JSON.stringify({ jsonrpc: '2.0', id: req.id, result: { tools: [{ name: 'slow_op' }] } }) + '\\n');
  } else if (req.method === 'tools/call') {
    // Hangs intentionally to test abort
  }
});
      `,
    );

    client = new StdioMcpClient({
      id: 'stdio-slow',
      name: 'Slow Stdio Server',
      enabled: true,
      transport: 'stdio',
      command: process.execPath,
      args: [serverScriptPath],
      allowedPaths: [],
    });

    await client.initialize();
    const controller = new AbortController();
    const callPromise = client.callTool('slow_op', {}, { abortSignal: controller.signal });
    controller.abort();

    await expect(callPromise).rejects.toThrow('Operation aborted');
  });
});
