import { createServer as createHttpServer } from 'node:http';
import type { AddressInfo } from 'node:net';
import { afterEach, describe, expect, it } from 'vitest';
import { WebSocket } from 'ws';
import { openDb } from '../src/db.js';
import { startServer, type SidecarServer } from '../src/server.js';

let sidecar: SidecarServer | null = null;

afterEach(async () => {
  await sidecar?.close();
  sidecar = null;
});

async function boot(): Promise<{ base: string; token: string }> {
  sidecar = await startServer({ db: openDb(':memory:'), port: 0 });
  const base = `http://127.0.0.1:${sidecar.port}`;
  const code = sidecar.auth.currentPairingCode!;
  const res = await fetch(`${base}/api/pair`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ code }),
  });
  const { token } = (await res.json()) as { token: string };
  return { base, token };
}

describe('sidecar HTTP API', () => {
  it('serves health without auth', async () => {
    sidecar = await startServer({ db: openDb(':memory:'), port: 0 });
    const res = await fetch(`http://127.0.0.1:${sidecar.port}/api/health`);
    expect(res.status).toBe(200);
    const body = (await res.json()) as { ok: boolean; protocol: number };
    expect(body.ok).toBe(true);
    expect(body.protocol).toBe(1);
  });

  it('rejects API calls without a valid token', async () => {
    const { base } = await boot();
    const res = await fetch(`${base}/api/ledger/summary`);
    expect(res.status).toBe(401);
    const forged = await fetch(`${base}/api/ledger/summary`, {
      headers: { authorization: 'Bearer forged' },
    });
    expect(forged.status).toBe(401);
  });

  it('allows DELETE in CORS preflights (model deletion from the webview)', async () => {
    sidecar = await startServer({ db: openDb(':memory:'), port: 0 });
    const res = await fetch(`http://127.0.0.1:${sidecar.port}/api/settings/providers/x`, {
      method: 'OPTIONS',
      headers: { origin: 'https://excel.officeapps.live.com' },
    });
    expect(res.status).toBe(204);
    expect(res.headers.get('access-control-allow-methods')).toContain('DELETE');
    expect(res.headers.get('access-control-allow-origin')).toBe('https://excel.officeapps.live.com');
  });

  it('pairs once only', async () => {
    const { base } = await boot();

    // pairing code is single-use
    const again = await fetch(`${base}/api/pair`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ code: 'ABCDEF' }),
    });
    expect(again.status).toBe(403);
  });

  it('manages memory entries over HTTP', async () => {
    const { base, token } = await boot();
    const auth = { authorization: `Bearer ${token}`, 'content-type': 'application/json' };
    const add = await fetch(`${base}/api/memory`, {
      method: 'POST',
      headers: auth,
      body: JSON.stringify({ scope: 'user', content: 'prefers ISO dates' }),
    });
    expect(add.status).toBe(200);
    const list = await fetch(`${base}/api/memory?scope=user`, { headers: auth });
    const body = (await list.json()) as { entries: { content: string }[] };
    expect(body.entries[0]?.content).toBe('prefers ISO dates');
  });

  it('handles abort message over WebSocket to stop in-flight generation', async () => {
    const { base, token } = await boot();
    const wsUrl = `${base.replace('http', 'ws')}/ws?token=${token}`;
    const ws = new WebSocket(wsUrl);

    const received: { type: string; payload: unknown }[] = [];
    ws.on('message', (data) => {
      received.push(JSON.parse(data.toString()));
    });

    await new Promise<void>((resolve) => ws.once('open', resolve));

    // Send client hello
    ws.send(
      JSON.stringify({
        v: 1,
        id: '1',
        type: 'hello',
        payload: { token, workbookId: 'wb1', workbookName: 'Wb.xlsx', client: 'addin' },
      }),
    );

    // Wait for hello_ack
    await new Promise<void>((resolve) => {
      const check = setInterval(() => {
        if (received.some((m) => m.type === 'hello_ack')) {
          clearInterval(check);
          resolve();
        }
      }, 10);
    });

    // Start a chat turn
    ws.send(
      JSON.stringify({
        v: 1,
        id: '2',
        type: 'chat',
        payload: { message: 'Calculate something complex' },
      }),
    );

    // Immediately send abort
    ws.send(
      JSON.stringify({
        v: 1,
        id: '3',
        type: 'abort',
        payload: { reason: 'user_requested' },
      }),
    );

    // Wait for done event
    await new Promise<void>((resolve) => {
      const check = setInterval(() => {
        const events = received
          .filter((m) => m.type === 'agent_event')
          .map((m) => m.payload as { kind: string });
        if (events.some((e) => e.kind === 'done')) {
          clearInterval(check);
          resolve();
        }
      }, 10);
    });

    const agentEvents = received
      .filter((m) => m.type === 'agent_event')
      .map((m) => m.payload as { kind: string });

    expect(agentEvents.some((e) => e.kind === 'stopped' || e.kind === 'error')).toBe(true);
    expect(agentEvents.some((e) => e.kind === 'done')).toBe(true);

    ws.close();
  });

  it('does not follow redirects during model discovery', async () => {
    const { base, token } = await boot();
    // Loopback endpoint whose /v1/models redirects to a second path that would succeed.
    const upstream = createHttpServer((req, res) => {
      if (req.url === '/v1/models') {
        res.writeHead(302, { location: '/elsewhere/models' }).end();
      } else {
        res.writeHead(200, { 'content-type': 'application/json' });
        res.end(JSON.stringify({ data: [{ id: 'redirected-model' }] }));
      }
    });
    await new Promise<void>((resolve) => upstream.listen(0, '127.0.0.1', resolve));
    try {
      const port = (upstream.address() as AddressInfo).port;
      const res = await fetch(`${base}/api/providers/models`, {
        method: 'POST',
        headers: { authorization: `Bearer ${token}`, 'content-type': 'application/json' },
        body: JSON.stringify({ kind: 'openai-compatible', baseUrl: `http://127.0.0.1:${port}/v1` }),
      });
      const body = (await res.json()) as { models: string[]; error?: string };
      expect(body.models).toEqual([]);
      expect(body.error).toBeDefined();
    } finally {
      await new Promise<void>((resolve) => upstream.close(() => resolve()));
    }
  });

  it('lists models and saves embedded provider configuration', async () => {
    const { base, token } = await boot();
    const auth = { authorization: `Bearer ${token}`, 'content-type': 'application/json' };

    // 1. List embedded models
    const listRes = await fetch(`${base}/api/providers/models`, {
      method: 'POST',
      headers: auth,
      body: JSON.stringify({ kind: 'embedded' }),
    });
    expect(listRes.status).toBe(200);
    const listData = (await listRes.json()) as { models: string[]; error?: string };
    expect(Array.isArray(listData.models)).toBe(true);

    // 2. Save embedded provider configuration
    const saveRes = await fetch(`${base}/api/settings/providers`, {
      method: 'POST',
      headers: auth,
      body: JSON.stringify({
        kind: 'embedded',
        model: 'qwen2.5-7b-instruct-q4_k_m.gguf',
        contextSize: 4096,
        threads: 4,
      }),
    });
    expect(saveRes.status).toBe(200);
    const saveData = (await saveRes.json()) as { id: string; activeProviderId: string };
    expect(saveData.id).toBeDefined();

    // 3. Verify in provider settings list
    const getRes = await fetch(`${base}/api/settings/providers`, { headers: auth });
    const getData = (await getRes.json()) as { providers: { kind: string; model: string }[] };
    const embedded = getData.providers.find((p) => p.model === 'qwen2.5-7b-instruct-q4_k_m.gguf');
    expect(embedded).toBeDefined();
    expect(embedded?.kind).toBe('embedded');
  });

  // The `anthropic` and `openai` kinds are wire protocols for local servers,
  // not cloud integrations: their vendors' URLs are refused everywhere.
  it('refuses cloud endpoints for the anthropic and openai protocol kinds', async () => {
    const { base, token } = await boot();
    const auth = { authorization: `Bearer ${token}`, 'content-type': 'application/json' };
    const cases = [
      { kind: 'anthropic', baseUrl: 'https://api.anthropic.com' },
      { kind: 'openai', baseUrl: 'https://api.openai.com/v1' },
    ];
    for (const c of cases) {
      for (const path of ['/api/settings/providers', '/api/settings/providers/test', '/api/providers/models']) {
        const res = await fetch(`${base}${path}`, {
          method: 'POST',
          headers: auth,
          body: JSON.stringify({ ...c, model: 'm', apiKey: 'k' }),
        });
        expect(res.status, `${c.kind} ${path}`).toBe(400);
        expect(JSON.stringify(await res.json())).toMatch(/loopback/);
      }
    }
    const list = (await (await fetch(`${base}/api/settings/providers`, { headers: auth })).json()) as {
      providers: { baseUrl?: string }[];
    };
    expect(list.providers.filter((p) => p.baseUrl && !p.baseUrl.startsWith('http://127.0.0.1'))).toEqual([]);

    const local = await fetch(`${base}/api/settings/providers`, {
      method: 'POST',
      headers: auth,
      body: JSON.stringify({ kind: 'anthropic', model: 'm', baseUrl: 'http://127.0.0.1:4000' }),
    });
    expect(local.status).toBe(200);
  });
});

describe('sidecar WebSocket hardening', () => {
  async function connect(base: string, token: string): Promise<WebSocket> {
    const ws = new WebSocket(`${base.replace('http', 'ws')}/ws?token=${token}`);
    await new Promise<void>((resolve, reject) => {
      ws.once('open', resolve);
      ws.once('error', reject);
    });
    return ws;
  }

  function closed(ws: WebSocket): Promise<number> {
    return new Promise((resolve) => ws.once('close', (code) => resolve(code)));
  }

  async function hello(ws: WebSocket, token: string, workbookId: string): Promise<{ ok: boolean }> {
    const ack = new Promise<{ ok: boolean }>((resolve) => {
      ws.on('message', (data) => {
        const msg = JSON.parse(data.toString()) as { type: string; payload: { ok: boolean } };
        if (msg.type === 'hello_ack') resolve(msg.payload);
      });
    });
    ws.send(
      JSON.stringify({
        v: 1,
        id: 'h',
        type: 'hello',
        payload: { token, workbookId, workbookName: `${workbookId}.xlsx`, client: 'addin' },
      }),
    );
    return ack;
  }

  type ProtocolErrorPayload = { message: string; issues: string[]; refId?: string };

  function protocolError(ws: WebSocket): Promise<ProtocolErrorPayload> {
    return new Promise((resolve) => {
      const onMessage = (data: unknown): void => {
        const msg = JSON.parse(String(data)) as { type: string; payload: ProtocolErrorPayload };
        if (msg.type !== 'protocol_error') return;
        ws.off('message', onMessage);
        resolve(msg.payload);
      };
      ws.on('message', onMessage);
    });
  }

  it.each([
    ['null', 'null'],
    ['an array', '[]'],
    ['a number', '42'],
    ['invalid JSON', '{"v":1,'],
    ['a frame missing its payload', JSON.stringify({ v: 1, id: 'x', type: 'hello' })],
    ['an unknown message type', JSON.stringify({ v: 1, id: 'x', type: 'drop_tables', payload: {} })],
    [
      'a hello with a wrongly typed field',
      JSON.stringify({ v: 1, id: 'x', type: 'hello', payload: { token: 't', workbookId: 5, workbookName: 'a', client: 'addin' } }),
    ],
    ['a chat whose message is not a string', JSON.stringify({ v: 1, id: 'x', type: 'chat', payload: { message: {} } })],
    ['a tool_result without ok', JSON.stringify({ v: 1, id: 'x', type: 'tool_result', payload: { requestId: 'r' } })],
    ['a frame without an id', JSON.stringify({ v: 1, type: 'abort', payload: {} })],
  ])('answers %s with a protocol error and keeps serving others', async (_label, frame) => {
    const { base, token } = await boot();
    const bystander = await connect(base, token);
    const attacker = await connect(base, token);

    const rejected = protocolError(attacker);
    attacker.send(frame);
    const error = await rejected;
    expect(error.message).toMatch(/malformed|JSON/i);
    expect(attacker.readyState).toBe(WebSocket.OPEN);

    // The offender can still speak the protocol correctly afterwards.
    expect((await hello(attacker, token, 'wb-attacker')).ok).toBe(true);
    expect((await hello(bystander, token, 'wb-bystander')).ok).toBe(true);
    const health = await fetch(`${base}/api/health`);
    expect(health.status).toBe(200);
    attacker.close();
    bystander.close();
  });

  it('names the offending field and echoes the frame id', async () => {
    const { base, token } = await boot();
    const ws = await connect(base, token);
    const rejected = protocolError(ws);
    ws.send(JSON.stringify({ v: 1, id: 'req-7', type: 'approve_changeset', payload: { changeSetId: 42 } }));
    const error = await rejected;
    expect(error.refId).toBe('req-7');
    expect(error.issues.some((i) => i.startsWith('payload.changeSetId'))).toBe(true);
    ws.close();
  });

  it('rejects binary frames', async () => {
    const { base, token } = await boot();
    const ws = await connect(base, token);
    const rejected = protocolError(ws);
    ws.send(Buffer.from(JSON.stringify({ v: 1, id: 'b', type: 'abort', payload: {} })));
    expect((await rejected).message).toMatch(/binary/i);
    ws.close();
  });

  it('closes a socket that keeps sending garbage while a second client and health stay responsive', async () => {
    const { base, token } = await boot();
    const bystander = await connect(base, token);
    const attacker = await connect(base, token);
    const errors: ProtocolErrorPayload[] = [];
    attacker.on('message', (data) => {
      const msg = JSON.parse(String(data)) as { type: string; payload: ProtocolErrorPayload };
      if (msg.type === 'protocol_error') errors.push(msg.payload);
    });

    const attackerClosed = closed(attacker);
    for (const frame of ['null', '[]', '"x"', '{}', 'nope']) {
      if (attacker.readyState === WebSocket.OPEN) attacker.send(frame);
    }
    const [code, health, ack] = await Promise.all([
      attackerClosed,
      fetch(`${base}/api/health`),
      hello(bystander, token, 'wb-bystander'),
    ]);
    expect(code).toBe(1008);
    expect(errors.length).toBe(3);
    expect(health.status).toBe(200);
    expect(ack.ok).toBe(true);
    bystander.close();
  });

  it('still answers a protocol version mismatch with a failed hello_ack', async () => {
    const { base, token } = await boot();
    const ws = await connect(base, token);
    const wsClosed = closed(ws);
    const ack = new Promise<{ ok: boolean; error?: string }>((resolve) => {
      ws.on('message', (data) => {
        const msg = JSON.parse(String(data)) as { type: string; payload: { ok: boolean; error?: string } };
        if (msg.type === 'hello_ack') resolve(msg.payload);
      });
    });
    ws.send(JSON.stringify({ v: 2, id: 'h', type: 'hello', payload: {} }));
    expect(await ack).toMatchObject({ ok: false, error: expect.stringMatching(/version mismatch/) });
    await wsClosed;
  });

  it('reports a change-set it no longer holds (e.g. after a restart) as unknown', async () => {
    const { base, token } = await boot();
    const ws = await connect(base, token);
    expect((await hello(ws, token, 'wb-restart')).ok).toBe(true);

    const status = new Promise<unknown>((resolve) => {
      ws.on('message', (data) => {
        const msg = JSON.parse(data.toString()) as { type: string; payload: { kind?: string } };
        if (msg.type === 'agent_event' && msg.payload.kind === 'changeset_status') resolve(msg.payload);
      });
    });
    ws.send(JSON.stringify({ v: 1, id: 'u', type: 'undo_changeset', payload: { changeSetId: 'gone' } }));

    expect(await status).toEqual({ kind: 'changeset_status', changeSetId: 'gone', status: 'unknown' });
    ws.close();
  });

  it('rejects frames larger than the payload cap', async () => {
    const { base, token } = await boot();
    const ws = await connect(base, token);
    ws.on('error', () => {});
    const wsClosed = closed(ws);
    ws.send('x'.repeat(16 * 1024 * 1024 + 1));
    expect(await wsClosed).toBe(1009);
  });
});

