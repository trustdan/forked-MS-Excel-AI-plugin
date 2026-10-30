import { mkdirSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import {
  clearSettingsProblem,
  dataDir,
  loadSettings,
  saveSettings,
  settingsProblem,
  SettingsUnavailableError,
  type Settings,
} from '../src/config.js';
import { openDb } from '../src/db.js';
import { startServer, type SidecarServer } from '../src/server.js';

const file = (): string => join(dataDir(), 'settings.json');
const settingsFiles = (): string[] => readdirSync(dataDir()).filter((f) => f.startsWith('settings.json'));

let sidecar: SidecarServer | null = null;

afterEach(async () => {
  await sidecar?.close();
  sidecar = null;
  for (const f of settingsFiles()) rmSync(join(dataDir(), f), { recursive: true, force: true });
  clearSettingsProblem();
});

const sample: Settings = {
  providers: [{ id: 'p1', kind: 'openai-compatible', model: 'local-model', baseUrl: 'http://127.0.0.1:1234/v1' }],
  activeProviderId: 'p1',
  mcpServers: [],
  approvedPaths: ['C:/data'],
};

describe('settings load (C5)', () => {
  it('treats a missing file as a fresh install, with no problem reported', () => {
    expect(loadSettings()).toEqual({ providers: [], activeProviderId: null, mcpServers: [], approvedPaths: [] });
    expect(settingsProblem()).toBeNull();
    expect(settingsFiles()).toEqual([]);
  });

  it('moves a file that is not JSON aside, byte for byte, and reports it', () => {
    writeFileSync(file(), '{"providers": [ truncated');
    expect(loadSettings().providers).toEqual([]);

    const problem = settingsProblem();
    expect(problem?.kind).toBe('unreadable');
    expect(problem?.preservedAs).toMatch(/settings\.json\.corrupt-/);
    expect(readFileSync(problem!.preservedAs, 'utf8')).toBe('{"providers": [ truncated');
    expect(settingsFiles()).toEqual([problem!.preservedAs.split(/[\\/]/).pop()]);
  });

  it('moves a file with the wrong shape aside and says which field is wrong', () => {
    writeFileSync(file(), JSON.stringify({ providers: 'not-a-list', approvedPaths: ['C:/keep-me'] }));
    expect(loadSettings().approvedPaths).toEqual([]);

    const problem = settingsProblem();
    expect(problem?.kind).toBe('invalid');
    expect(problem?.detail).toContain('providers');
    expect(JSON.parse(readFileSync(problem!.preservedAs, 'utf8'))).toEqual({
      providers: 'not-a-list',
      approvedPaths: ['C:/keep-me'],
    });
  });

  it('a save after a corrupt load does not overwrite the preserved original', () => {
    writeFileSync(file(), 'garbage');
    const settings = loadSettings();
    settings.approvedPaths.push('C:/new');
    saveSettings(settings);
    expect(readFileSync(settingsProblem()!.preservedAs, 'utf8')).toBe('garbage');
    expect(loadSettings().approvedPaths).toEqual(['C:/new']);
  });

  it('refuses, and changes nothing, when the file exists but cannot be read', () => {
    mkdirSync(file()); // reading a directory fails with EISDIR, not ENOENT
    expect(() => loadSettings()).toThrow(SettingsUnavailableError);
    expect(statSync(file()).isDirectory()).toBe(true);
    expect(settingsProblem()).toBeNull();
    expect(settingsFiles()).toEqual(['settings.json']);
  });

  it('keeps fields it does not know about across a load and save', () => {
    writeFileSync(file(), JSON.stringify({ ...sample, providers: [{ ...sample.providers[0], futureField: 7 }] }));
    saveSettings(loadSettings());
    const saved = JSON.parse(readFileSync(file(), 'utf8')) as { providers: { futureField?: number }[] };
    expect(saved.providers[0]?.futureField).toBe(7);
  });

  it('still migrates the legacy single-provider shape', () => {
    writeFileSync(file(), JSON.stringify({ provider: { kind: 'openai-compatible', model: 'old' } }));
    const settings = loadSettings();
    expect(settings.providers).toHaveLength(1);
    expect(settings.providers[0]?.model).toBe('old');
    expect(settings.activeProviderId).toBe(settings.providers[0]?.id);
    expect(settingsProblem()).toBeNull();
  });
});

describe('settings save (C5)', () => {
  it('replaces the file whole and leaves no temp files behind', () => {
    saveSettings(sample);
    saveSettings({ ...sample, approvedPaths: ['C:/second'] });
    expect(settingsFiles()).toEqual(['settings.json']);
    expect(loadSettings().approvedPaths).toEqual(['C:/second']);
  });

  it('on a failed replace, keeps the old file intact and cleans up the temp file', () => {
    // A non-empty directory at the target makes the final rename fail on every platform.
    mkdirSync(join(file(), 'blocker'), { recursive: true });
    expect(() => saveSettings(sample)).toThrow();
    expect(statSync(file()).isDirectory()).toBe(true);
    expect(settingsFiles()).toEqual(['settings.json']);
  });
});

describe('settings problem in the taskpane API (C5)', () => {
  it('reports the problem on the models list until dismissed', async () => {
    writeFileSync(file(), 'not json');
    sidecar = await startServer({ db: openDb(':memory:'), port: 0 });
    const base = `http://127.0.0.1:${sidecar.port}`;
    const pair = await fetch(`${base}/api/pair`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ code: sidecar.auth.currentPairingCode }),
    });
    const { token } = (await pair.json()) as { token: string };
    const headers = { authorization: `Bearer ${token}` };

    const list = async (): Promise<{ settingsProblem?: { kind: string; preservedAs: string } }> =>
      (await fetch(`${base}/api/settings/providers`, { headers })).json() as Promise<{
        settingsProblem?: { kind: string; preservedAs: string };
      }>;
    const before = await list();
    expect(before.settingsProblem?.kind).toBe('unreadable');
    expect(before.settingsProblem?.preservedAs).toMatch(/settings\.json\.corrupt-/);

    expect((await fetch(`${base}/api/settings/problem`, { method: 'DELETE', headers })).status).toBe(200);
    expect((await list()).settingsProblem).toBeUndefined();
  });

  it('starts, and reports the error on settings routes, when the file cannot be read', async () => {
    mkdirSync(file());
    sidecar = await startServer({ db: openDb(':memory:'), port: 0 });
    const base = `http://127.0.0.1:${sidecar.port}`;
    expect((await fetch(`${base}/api/health`)).status).toBe(200);
    const pair = await fetch(`${base}/api/pair`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ code: sidecar.auth.currentPairingCode }),
    });
    const { token } = (await pair.json()) as { token: string };
    const res = await fetch(`${base}/api/settings/providers`, { headers: { authorization: `Bearer ${token}` } });
    expect(res.status).toBe(500);
    expect(((await res.json()) as { error: string }).error).toContain('refusing to continue with empty settings');
  });
});
