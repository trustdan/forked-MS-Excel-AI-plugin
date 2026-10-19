import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { mkdirSync, writeFileSync, rmSync } from 'node:fs';
import { McpManager } from '../src/mcp/manager.js';
import type { McpServerConfig } from '@excelai/shared';

describe('McpManager', () => {
  let tempDir: string;
  let manager: McpManager;

  beforeEach(() => {
    tempDir = join(tmpdir(), `excelai-manager-test-${Date.now()}`);
    mkdirSync(tempDir, { recursive: true });
    writeFileSync(join(tempDir, 'sample.txt'), 'Manager test data');
    manager = new McpManager();
  });

  afterEach(async () => {
    await manager.closeAll();
    try {
      rmSync(tempDir, { recursive: true, force: true });
    } catch {}
  });

  it('syncs servers and registers tools with untrusted data wrappers', async () => {
    const config: McpServerConfig = {
      id: 'fs-server',
      name: 'TestFS',
      enabled: true,
      transport: 'builtin',
      type: 'filesystem',
      allowedPaths: [tempDir],
      options: { rootPath: tempDir },
    };

    await manager.sync([config], [tempDir]);
    expect(manager.hasTool('fs_read_file')).toBe(true);

    const specs = manager.getToolSpecs();
    expect(specs.some((s) => s.name === 'fs_read_file')).toBe(true);

    const callRes = await manager.callTool('fs_read_file', { path: 'sample.txt' });
    expect(callRes.isError).toBeFalsy();
    // Prompt injection defense check: Untrusted marker wrapper
    expect(callRes.resultText).toContain('[UNTRUSTED EXTERNAL DATA FROM MCP SERVER: TestFS]');
    expect(callRes.resultText).toContain('Manager test data');
    expect(callRes.resultText).toContain('[/UNTRUSTED EXTERNAL DATA]');
  });

  it('tests a candidate server configuration without persisting', async () => {
    const config: McpServerConfig = {
      id: 'candidate-fs',
      name: 'CandidateFS',
      enabled: true,
      transport: 'builtin',
      type: 'filesystem',
      allowedPaths: [tempDir],
      options: { rootPath: tempDir },
    };

    const res = await manager.testServer(config, [tempDir]);
    expect(res.ok).toBe(true);
    expect(res.tools.length).toBeGreaterThan(0);
    expect(res.tools.map((t) => t.name)).toContain('fs_list_directory');
  });

  it('unregisters tools when server is disabled', async () => {
    const config: McpServerConfig = {
      id: 'fs-server',
      name: 'TestFS',
      enabled: true,
      transport: 'builtin',
      type: 'filesystem',
      allowedPaths: [tempDir],
      options: { rootPath: tempDir },
    };

    await manager.sync([config], [tempDir]);
    expect(manager.hasTool('fs_read_file')).toBe(true);

    // Disable server
    await manager.sync([{ ...config, enabled: false }], [tempDir]);
    expect(manager.hasTool('fs_read_file')).toBe(false);
  });
});
