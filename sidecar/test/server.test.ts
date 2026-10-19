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
});

