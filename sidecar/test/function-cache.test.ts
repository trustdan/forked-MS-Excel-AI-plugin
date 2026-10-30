import { createServer, type Server } from 'node:http';
import type { AddressInfo } from 'node:net';
import { rmSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { dataDir, providerIdentity, type StoredProvider } from '../src/config.js';
import { openDb } from '../src/db.js';
import { startServer, type SidecarServer } from '../src/server.js';

let sidecar: SidecarServer | null = null;
const models: Server[] = [];

afterEach(async () => {
  await sidecar?.close();
  sidecar = null;
  await Promise.all(models.splice(0).map((m) => new Promise<void>((r) => m.close(() => r()))));
  rmSync(join(dataDir(), 'settings.json'), { force: true });
});

/** A loopback OpenAI-compatible server that always answers `reply` and counts calls. */
async function mockModel(reply: string): Promise<{ baseUrl: string; calls: () => number }> {
  let calls = 0;
  const server = createServer((req, res) => {
    req.resume();
    req.on('end', () => {
      calls++;
      res.writeHead(200, { 'content-type': 'application/json' });
      res.end(
        JSON.stringify({
          choices: [{ message: { content: reply }, finish_reason: 'stop' }],
          usage: { prompt_tokens: 1, completion_tokens: 1 },
        }),
      );
    });
  });
  models.push(server);
  await new Promise<void>((r) => server.listen(0, '127.0.0.1', r));
  return { baseUrl: `http://127.0.0.1:${(server.address() as AddressInfo).port}/v1`, calls: () => calls };
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

describe('function cache identity (F12)', () => {
  it('switching between two configs with the same model name uses the selected one', async () => {
    const a = await mockModel('server A');
    const b = await mockModel('server B');
    const { base, headers } = await boot();
    const save = async (baseUrl: string): Promise<string> => {
      const res = await fetch(`${base}/api/settings/providers`, {
        method: 'POST',
        headers,
        body: JSON.stringify({ kind: 'openai-compatible', model: 'local-model', baseUrl }),
      });
      return ((await res.json()) as { id: string }).id;
    };
    const idA = await save(a.baseUrl);
    const idB = await save(b.baseUrl);
    const activate = (id: string): Promise<Response> =>
      fetch(`${base}/api/settings/providers/${id}/activate`, { method: 'POST', headers });
    const run = async (): Promise<string> => {
      const res = await fetch(`${base}/api/functions/run`, {
        method: 'POST',
        headers,
        body: JSON.stringify({ kind: 'ai', prompt: 'who are you' }),
      });
      expect(res.status).toBe(200);
      return ((await res.json()) as { value: string }).value;
    };

    await activate(idA);
    expect(await run()).toBe('server A');
    await activate(idB);
    expect(await run()).toBe('server B');
    expect(b.calls()).toBe(1);

    // Unchanged configuration: repeats still come from the cache.
    expect(await run()).toBe('server B');
    expect(b.calls()).toBe(1);
    await activate(idA);
    expect(await run()).toBe('server A');
    expect(a.calls()).toBe(1);
  });
});

describe('providerIdentity', () => {
  const base: StoredProvider = {
    id: 'cfg-1',
    kind: 'openai-compatible',
    model: 'local-model',
    baseUrl: 'http://127.0.0.1:1234/v1',
    apiKey: 'secret-key',
  };

  it('changes with the config id, endpoint, model or model file', () => {
    const id = providerIdentity(base);
    expect(providerIdentity({ ...base, id: 'cfg-2' })).not.toBe(id);
    expect(providerIdentity({ ...base, baseUrl: 'http://127.0.0.1:11434/v1' })).not.toBe(id);
    expect(providerIdentity({ ...base, model: 'other-model' })).not.toBe(id);
    expect(providerIdentity({ ...base, modelPath: 'C:/models/a.gguf' })).not.toBe(id);
  });

  it('does not depend on, or contain, the API key', () => {
    const id = providerIdentity(base);
    expect(providerIdentity({ ...base, apiKey: 'rotated-key' })).toBe(id);
    expect(id).not.toContain('secret-key');
  });

  it('does not depend on key order', () => {
    const reordered = JSON.parse(JSON.stringify({ apiKey: 'x', model: base.model, baseUrl: base.baseUrl, kind: base.kind, id: base.id })) as StoredProvider;
    expect(providerIdentity(reordered)).toBe(providerIdentity(base));
  });
});
