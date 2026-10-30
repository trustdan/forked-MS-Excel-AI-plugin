import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { FilesystemConnector } from '../src/mcp/connectors/filesystemConnector.js';

// Its own file so the forked test process has no earlier peak to hide behind.
describe('FilesystemConnector memory bound (C4)', () => {
  let rootDir: string;

  beforeEach(() => {
    rootDir = join(tmpdir(), `excelai-fs-mem-${crypto.randomUUID()}`);
    mkdirSync(rootDir, { recursive: true });
  });

  afterEach(() => {
    rmSync(rootDir, { recursive: true, force: true });
  });

  it('reads a 50 MB file without allocating it', async () => {
    const chunk = Buffer.alloc(1024 * 1024, 'a');
    for (let i = 0; i < 50; i++) writeFileSync(join(rootDir, 'big.txt'), chunk, { flag: 'a' });

    const connector = new FilesystemConnector({
      id: 'fs-1',
      name: 'Local Workspace',
      enabled: true,
      transport: 'builtin',
      type: 'filesystem',
      allowedPaths: [rootDir],
      options: { rootPath: rootDir },
    });

    const peakBefore = process.resourceUsage().maxRSS; // KiB
    const res = await connector.callTool('fs_read_file', { path: 'big.txt', maxBytes: 1024 });
    const peakGrowthMiB = (process.resourceUsage().maxRSS - peakBefore) / 1024;

    expect(JSON.parse(res.content[0]!.text)).toMatchObject({ bytesRead: 1024, truncated: true });
    expect(peakGrowthMiB).toBeLessThan(20);
  });
});
