import { createServer, type IncomingMessage, type ServerResponse } from 'node:http';
import type { AddressInfo } from 'node:net';
import { afterEach, describe, expect, it } from 'vitest';
import { openDb } from '../src/db.js';
import {
  guardedFetch,
  isLoopbackAddress,
  NonLocalConnectionBlockedError,
  RedirectBlockedError,
} from '../src/networkGuard.js';
import { AnthropicProvider } from '../src/providers/anthropic.js';
import { OpenAIProvider } from '../src/providers/openai.js';
import { startServer, type SidecarServer } from '../src/server.js';

// Every test here goes through the production transport against real loopback sockets.

interface Fixture {
  port: number;
  hits: number;
  /** Resolves when the server sees the client's connection close. */
  closed: Promise<void>;
}

const cleanups: (() => Promise<void>)[] = [];
afterEach(async () => {
  while (cleanups.length) await cleanups.pop()!();
});

async function serve(
  handler: (req: IncomingMessage, res: ServerResponse) => void,
  host = '127.0.0.1',
): Promise<Fixture> {
  let markClosed!: () => void;
  const fixture: Fixture = { port: 0, hits: 0, closed: new Promise((r) => (markClosed = r)) };
  const server = createServer((req, res) => {
    fixture.hits++;
    res.on('close', markClosed);
    handler(req, res);
  });
  await new Promise<void>((r) => server.listen(0, host, r));
  fixture.port = (server.address() as AddressInfo).port;
  cleanups.push(async () => {
    server.closeAllConnections();
    await new Promise<void>((r) => server.close(() => r()));
  });
  return fixture;
}

const sseHeaders = { 'content-type': 'text/event-stream' };

function within<T>(promise: Promise<T>, ms = 2_000): Promise<T> {
  return Promise.race([
    promise,
    new Promise<never>((_, reject) => setTimeout(() => reject(new Error(`not settled within ${ms}ms`)), ms)),
  ]);
}

describe('guardedFetch streaming', () => {
  it('delivers the first SSE delta to the caller before the server ends the stream', async () => {
    let release!: () => void;
    const released = new Promise<void>((r) => (release = r));
    let serverEnded = false;
    const model = await serve((_req, res) => {
      res.writeHead(200, sseHeaders);
      res.write(`data: ${JSON.stringify({ choices: [{ delta: { content: 'Hello' } }] })}\n\n`);
      void released.then(() => {
        res.write(`data: ${JSON.stringify({ choices: [{ delta: {}, finish_reason: 'stop' }] })}\n\n`);
        res.end('data: [DONE]\n\n');
        serverEnded = true;
      });
    });

    const provider = new OpenAIProvider('m', 'k', `http://127.0.0.1:${model.port}/v1`);
    const deltas: string[] = [];
    let endedAtFirstDelta: boolean | undefined;
    const result = await within(
      provider.chat({
        messages: [{ role: 'user', content: 'hi' }],
        onDelta: (d) => {
          if (!d.text) return;
          endedAtFirstDelta ??= serverEnded;
          deltas.push(d.text);
          release();
        },
      }),
    );

    expect(endedAtFirstDelta).toBe(false);
    expect(deltas).toEqual(['Hello']);
    expect(result.text).toBe('Hello');
  });

  it('resolves `localhost` to a loopback address and connects to it', async () => {
    const model = await serve((_req, res) => res.end('ok'));
    const res = await guardedFetch(`http://localhost:${model.port}/`);
    expect(await res.text()).toBe('ok');
  });

  it('times out a stalled stream while the connection is still open', async () => {
    const model = await serve((_req, res) => {
      res.writeHead(200, sseHeaders);
      res.write('data: first\n\n'); // then stall forever
    });
    const res = await guardedFetch(`http://127.0.0.1:${model.port}/`, { inactivityTimeoutMs: 150 });
    const reader = res.body!.getReader();
    expect((await reader.read()).done).toBe(false);
    await expect(within(reader.read())).rejects.toMatchObject({ status: 504 });
    await within(model.closed);
  });

  it('rejects an oversized streamed body before retaining it, and drops the connection', async () => {
    const total = 4 * 1024 * 1024;
    let written = 0;
    const model = await serve((_req, res) => {
      res.writeHead(200, { 'content-type': 'application/octet-stream' });
      const pump = (): void => {
        while (written < total && !res.destroyed) {
          written += 16 * 1024;
          if (!res.write(Buffer.alloc(16 * 1024))) return void res.once('drain', pump);
        }
        res.end();
      };
      pump();
    });

    const res = await guardedFetch(`http://127.0.0.1:${model.port}/`, { maxResponseBytes: 64 * 1024 });
    await expect(within(res.arrayBuffer())).rejects.toMatchObject({ status: 413 });
    await within(model.closed);
    expect(written).toBeLessThan(total);
  });

  it('rejects a declared Content-Length over the cap before reading the body', async () => {
    const model = await serve((_req, res) => {
      res.writeHead(200, { 'content-length': String(10 * 1024 * 1024) });
      res.write(Buffer.alloc(1024));
    });
    await expect(
      guardedFetch(`http://127.0.0.1:${model.port}/`, { maxResponseBytes: 1024 * 1024 }),
    ).rejects.toMatchObject({ status: 413 });
    await within(model.closed);
  });

  it('closes the socket when the reader cancels', async () => {
    const model = await serve((_req, res) => {
      res.writeHead(200, sseHeaders);
      const timer = setInterval(() => res.write('data: tick\n\n'), 20);
      res.on('close', () => clearInterval(timer));
    });
    const res = await guardedFetch(`http://127.0.0.1:${model.port}/`);
    const reader = res.body!.getReader();
    await reader.read();
    await reader.cancel();
    await within(model.closed);
  });

  it('closes the socket when the caller aborts mid-stream', async () => {
    const model = await serve((_req, res) => {
      res.writeHead(200, sseHeaders);
      const timer = setInterval(() => res.write('data: tick\n\n'), 20);
      res.on('close', () => clearInterval(timer));
    });
    const controller = new AbortController();
    const res = await guardedFetch(`http://127.0.0.1:${model.port}/`, { signal: controller.signal });
    const reader = res.body!.getReader();
    await reader.read();
    controller.abort();
    await expect(within(reader.read())).rejects.toMatchObject({ name: 'AbortError' });
    await within(model.closed);
  });
});

describe('guardedFetch address and redirect policy', () => {
  it('accepts only 127.0.0.0/8 and ::1 as connected addresses', () => {
    for (const a of ['127.0.0.1', '127.8.9.10', '::1', '0:0:0:0:0:0:0:1']) expect(isLoopbackAddress(a)).toBe(true);
    for (const a of ['::ffff:127.0.0.1', '::ffff:7f00:1', '10.0.0.1', '0.0.0.0', '::', 'localhost']) {
      expect(isLoopbackAddress(a)).toBe(false);
    }
  });

  it('refuses an IPv4-mapped IPv6 URL before connecting', async () => {
    const target = await serve((_req, res) => res.end('reached'), '::');
    await expect(guardedFetch(`http://[::ffff:127.0.0.1]:${target.port}/`)).rejects.toBeInstanceOf(
      NonLocalConnectionBlockedError,
    );
    expect(target.hits).toBe(0);
  });

  it('refuses a redirect to an IPv4-mapped IPv6 address without following it', async () => {
    const target = await serve((_req, res) => res.end('reached'), '::');
    const origin = await serve((_req, res) => {
      res.writeHead(302, { location: `http://[::ffff:127.0.0.1]:${target.port}/` }).end();
    });
    await expect(guardedFetch(`http://127.0.0.1:${origin.port}/`)).rejects.toBeInstanceOf(RedirectBlockedError);
    expect(target.hits).toBe(0);
  });
});

describe('every outbound call site refuses redirects', () => {
  let sidecar: SidecarServer | null = null;
  afterEach(async () => {
    await sidecar?.close();
    sidecar = null;
  });

  /** A model endpoint that redirects every request to a second server that would answer successfully. */
  async function redirectingModel(): Promise<{ baseUrl: string; target: Fixture }> {
    const okChat = { choices: [{ message: { content: 'reached' }, finish_reason: 'stop' }], data: [{ id: 'm' }] };
    const target = await serve((_req, res) => {
      res.writeHead(200, { 'content-type': 'application/json' }).end(JSON.stringify(okChat));
    }, '::');
    const origin = await serve((req, res) => {
      res.writeHead(307, { location: `http://[::ffff:127.0.0.1]:${target.port}${req.url}` }).end();
    });
    return { baseUrl: `http://127.0.0.1:${origin.port}/v1`, target };
  }

  async function boot(): Promise<{ base: string; headers: Record<string, string> }> {
    sidecar = await startServer({ db: openDb(':memory:'), port: 0 });
    const base = `http://127.0.0.1:${sidecar.port}`;
    const res = await fetch(`${base}/api/pair`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ code: sidecar.auth.currentPairingCode }),
    });
    const { token } = (await res.json()) as { token: string };
    return { base, headers: { authorization: `Bearer ${token}`, 'content-type': 'application/json' } };
  }

  it('chat (OpenAI-compatible and Anthropic adapters)', async () => {
    const { baseUrl, target } = await redirectingModel();
    const messages = [{ role: 'user' as const, content: 'hi' }];
    await expect(new OpenAIProvider('m', 'k', baseUrl).chat({ messages })).rejects.toBeInstanceOf(
      RedirectBlockedError,
    );
    await expect(new AnthropicProvider('m', 'k', baseUrl).chat({ messages })).rejects.toBeInstanceOf(
      RedirectBlockedError,
    );
    expect(target.hits).toBe(0);
  });

  it('model discovery', async () => {
    const { baseUrl, target } = await redirectingModel();
    const { base, headers } = await boot();
    for (const kind of ['openai-compatible', 'anthropic']) {
      const res = await fetch(`${base}/api/providers/models`, {
        method: 'POST',
        headers,
        body: JSON.stringify({ kind, baseUrl }),
      });
      const body = (await res.json()) as { models: string[]; error?: string };
      expect(body.models).toEqual([]);
      expect(body.error).toMatch(/Redirect .* blocked/);
    }
    expect(target.hits).toBe(0);
  });

  it('provider probes', async () => {
    const { baseUrl, target } = await redirectingModel();
    const { base, headers } = await boot();
    const res = await fetch(`${base}/api/settings/providers/test`, {
      method: 'POST',
      headers,
      body: JSON.stringify({ kind: 'openai-compatible', model: 'm', baseUrl }),
    });
    const body = (await res.json()) as { ok: boolean; error?: string };
    expect(body.ok).toBe(false);
    expect(body.error).toMatch(/Redirect .* blocked/);
    expect(target.hits).toBe(0);
  });

  it('custom functions', async () => {
    const { baseUrl, target } = await redirectingModel();
    const { base, headers } = await boot();
    const saved = await fetch(`${base}/api/settings/providers`, {
      method: 'POST',
      headers,
      body: JSON.stringify({ name: 'redirecting', kind: 'openai-compatible', model: 'm', baseUrl }),
    });
    expect(saved.ok).toBe(true);
    const res = await fetch(`${base}/api/functions/run`, {
      method: 'POST',
      headers,
      body: JSON.stringify({ kind: 'ai', prompt: 'hi' }),
    });
    expect(res.status).toBe(502);
    expect(((await res.json()) as { error: string }).error).toMatch(/Redirect .* blocked/);
    expect(target.hits).toBe(0);
  });
});
