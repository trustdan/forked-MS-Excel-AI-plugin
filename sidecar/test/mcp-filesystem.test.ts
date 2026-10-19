import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { mkdirSync, writeFileSync, rmSync } from 'node:fs';
import { FilesystemConnector } from '../src/mcp/connectors/filesystemConnector.js';
import { McpSecurityError } from '../src/mcp/pathGuard.js';

describe('FilesystemConnector', () => {
  let tempDir: string;
  let rootDir: string;
  let connector: FilesystemConnector | null = null;

  beforeEach(() => {
    tempDir = join(tmpdir(), `excelai-fs-test-${Date.now()}`);
    rootDir = join(tempDir, 'workspace');
    mkdirSync(rootDir, { recursive: true });

    writeFileSync(join(rootDir, 'readme.txt'), 'Hello world from local fs!');
    const sub = join(rootDir, 'src');
    mkdirSync(sub, { recursive: true });
    writeFileSync(join(sub, 'index.ts'), 'console.log("ready");');

    // Binary file containing null byte
    writeFileSync(join(rootDir, 'binary.bin'), Buffer.from([0x00, 0x01, 0x02, 0x03]));
  });

  afterEach(async () => {
    if (connector) {
      await connector.close();
      connector = null;
    }
    try {
      rmSync(tempDir, { recursive: true, force: true });
    } catch {}
  });

  it('rejects root path outside approved paths', async () => {
    connector = new FilesystemConnector({
      id: 'fs-1',
      name: 'Local Workspace',
      enabled: true,
      transport: 'builtin',
      type: 'filesystem',
      allowedPaths: [join(tempDir, 'other')],
      options: { rootPath: rootDir },
    });

    await expect(connector.initialize()).rejects.toThrow(McpSecurityError);
  });

  it('lists directory contents', async () => {
    connector = new FilesystemConnector({
      id: 'fs-1',
      name: 'Local Workspace',
      enabled: true,
      transport: 'builtin',
      type: 'filesystem',
      allowedPaths: [rootDir],
      options: { rootPath: rootDir },
    });

    await connector.initialize();
    const res = await connector.callTool('fs_list_directory', {});
    expect(res.isError).toBeFalsy();
    const data = JSON.parse(res.content[0].text);
    expect(data.ok).toBe(true);
    const names = data.items.map((i: any) => i.name);
    expect(names).toContain('readme.txt');
    expect(names).toContain('src');
  });

  it('reads text files and rejects binary files', async () => {
    connector = new FilesystemConnector({
      id: 'fs-1',
      name: 'Local Workspace',
      enabled: true,
      transport: 'builtin',
      type: 'filesystem',
      allowedPaths: [rootDir],
      options: { rootPath: rootDir },
    });

    const textRes = await connector.callTool('fs_read_file', { path: 'readme.txt' });
    expect(textRes.isError).toBeFalsy();
    const textData = JSON.parse(textRes.content[0].text);
    expect(textData.content).toBe('Hello world from local fs!');

    const binRes = await connector.callTool('fs_read_file', { path: 'binary.bin' });
    expect(binRes.isError).toBe(true);
  });

  it('searches files by query', async () => {
    connector = new FilesystemConnector({
      id: 'fs-1',
      name: 'Local Workspace',
      enabled: true,
      transport: 'builtin',
      type: 'filesystem',
      allowedPaths: [rootDir],
      options: { rootPath: rootDir },
    });

    const res = await connector.callTool('fs_search_files', { query: 'index' });
    expect(res.isError).toBeFalsy();
    const data = JSON.parse(res.content[0].text);
    expect(data.totalMatches).toBe(1);
    expect(data.matches[0].relativePath).toBe('src/index.ts');
  });

  it('gets file metadata', async () => {
    connector = new FilesystemConnector({
      id: 'fs-1',
      name: 'Local Workspace',
      enabled: true,
      transport: 'builtin',
      type: 'filesystem',
      allowedPaths: [rootDir],
      options: { rootPath: rootDir },
    });

    const res = await connector.callTool('fs_get_metadata', { path: 'readme.txt' });
    expect(res.isError).toBeFalsy();
    const data = JSON.parse(res.content[0].text);
    expect(data.isFile).toBe(true);
    expect(data.sizeBytes).toBeGreaterThan(0);
  });

  it('rejects path traversal in fs_read_file', async () => {
    connector = new FilesystemConnector({
      id: 'fs-1',
      name: 'Local Workspace',
      enabled: true,
      transport: 'builtin',
      type: 'filesystem',
      allowedPaths: [rootDir],
      options: { rootPath: rootDir },
    });

    await expect(
      connector.callTool('fs_read_file', { path: '../outside.txt' }),
    ).rejects.toThrow(McpSecurityError);
  });
});
