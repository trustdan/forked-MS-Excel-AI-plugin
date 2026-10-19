import { afterEach, describe, expect, it } from 'vitest';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { openDb } from '../src/db.js';
import { startServer, type SidecarServer } from '../src/server.js';

let sidecar: SidecarServer | null = null;
let tempDir: string | null = null;

afterEach(async () => {
  await sidecar?.close();
  sidecar = null;
  if (tempDir) {
    try {
      rmSync(tempDir, { recursive: true, force: true });
    } catch {}
    tempDir = null;
  }
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

describe('MCP HTTP Endpoints', () => {
  it('manages approved paths (GET, POST, DELETE)', async () => {
    const { base, token } = await boot();
    tempDir = join(tmpdir(), `excelai-mcp-path-test-${Date.now()}`);
    mkdirSync(tempDir, { recursive: true });

    // GET empty initially or default
    const getRes = await fetch(`${base}/api/mcp/paths`, {
      headers: { authorization: `Bearer ${token}` },
    });
    expect(getRes.status).toBe(200);
    const getBody = (await getRes.json()) as { paths: string[] };
    expect(Array.isArray(getBody.paths)).toBe(true);

    // POST add path
    const addRes = await fetch(`${base}/api/mcp/paths`, {
      method: 'POST',
      headers: {
        authorization: `Bearer ${token}`,
        'content-type': 'application/json',
      },
      body: JSON.stringify({ path: tempDir }),
    });
    expect(addRes.status).toBe(200);
    const addBody = (await addRes.json()) as { ok: boolean; paths: string[] };
    expect(addBody.ok).toBe(true);

    // DELETE remove path
    const delRes = await fetch(`${base}/api/mcp/paths`, {
      method: 'DELETE',
      headers: {
        authorization: `Bearer ${token}`,
        'content-type': 'application/json',
      },
      body: JSON.stringify({ path: tempDir }),
    });
    expect(delRes.status).toBe(200);
    const delBody = (await delRes.json()) as { ok: boolean; paths: string[] };
    expect(delBody.ok).toBe(true);
  });

  it('manages MCP servers (list, add, test, toggle, delete)', async () => {
    const { base, token } = await boot();
    tempDir = join(tmpdir(), `excelai-mcp-srv-test-${Date.now()}`);
    mkdirSync(tempDir, { recursive: true });
    writeFileSync(join(tempDir, 'test.csv'), 'id,name\n1,Alpha\n');

    // 1. Add approved path
    await fetch(`${base}/api/mcp/paths`, {
      method: 'POST',
      headers: {
        authorization: `Bearer ${token}`,
        'content-type': 'application/json',
      },
      body: JSON.stringify({ path: tempDir }),
    });

    // 2. Test candidate server
    const testRes = await fetch(`${base}/api/mcp/servers/test`, {
      method: 'POST',
      headers: {
        authorization: `Bearer ${token}`,
        'content-type': 'application/json',
      },
      body: JSON.stringify({
        name: 'Candidate CSV',
        transport: 'builtin',
        type: 'csv-directory',
        allowedPaths: [tempDir],
        options: { dirPath: tempDir },
      }),
    });
    expect(testRes.status).toBe(200);
    const testBody = (await testRes.json()) as { ok: boolean; tools: any[] };
    expect(testBody.ok).toBe(true);
    expect(testBody.tools.map((t) => t.name)).toContain('csv_list_files');

    // 3. Add MCP server
    const addRes = await fetch(`${base}/api/mcp/servers`, {
      method: 'POST',
      headers: {
        authorization: `Bearer ${token}`,
        'content-type': 'application/json',
      },
      body: JSON.stringify({
        name: 'Live CSV Connector',
        enabled: true,
        transport: 'builtin',
        type: 'csv-directory',
        allowedPaths: [tempDir],
        options: { dirPath: tempDir },
      }),
    });
    expect(addRes.status).toBe(200);
    const addBody = (await addRes.json()) as { ok: boolean; server: { id: string; name: string } };
    expect(addBody.ok).toBe(true);
    const serverId = addBody.server.id;

    // 4. List servers
    const listRes = await fetch(`${base}/api/mcp/servers`, {
      headers: { authorization: `Bearer ${token}` },
    });
    expect(listRes.status).toBe(200);
    const listBody = (await listRes.json()) as { servers: any[] };
    expect(listBody.servers.some((s) => s.id === serverId)).toBe(true);

    // 5. Query active MCP tools
    const toolsRes = await fetch(`${base}/api/mcp/tools`, {
      headers: { authorization: `Bearer ${token}` },
    });
    expect(toolsRes.status).toBe(200);
    const toolsBody = (await toolsRes.json()) as { tools: any[] };
    expect(toolsBody.tools.some((t) => t.name === 'csv_list_files')).toBe(true);

    // 6. Toggle server
    const toggleRes = await fetch(`${base}/api/mcp/servers/${serverId}/toggle`, {
      method: 'POST',
      headers: { authorization: `Bearer ${token}` },
    });
    expect(toggleRes.status).toBe(200);
    const toggleBody = (await toggleRes.json()) as { ok: boolean; enabled: boolean };
    expect(toggleBody.enabled).toBe(false);

    // 7. Delete server
    const delRes = await fetch(`${base}/api/mcp/servers/${serverId}`, {
      method: 'DELETE',
      headers: { authorization: `Bearer ${token}` },
    });
    expect(delRes.status).toBe(200);
  });
});
