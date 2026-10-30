import { closeSync, fsyncSync, mkdirSync, openSync, readFileSync, renameSync, rmSync, writeFileSync } from 'node:fs';
import { createHash, randomUUID } from 'node:crypto';
import { homedir } from 'node:os';
import { dirname, join } from 'node:path';
import { z } from 'zod';
import type { ProviderConfig } from './providers/types.js';
import type { McpServerConfig } from '@excelai/shared';

/** A saved model configuration the user can switch between. */
export interface StoredProvider extends ProviderConfig {
  id: string;
}

export interface Settings {
  providers: StoredProvider[];
  activeProviderId: string | null;
  mcpServers: McpServerConfig[];
  approvedPaths: string[];
}

// Structural check only. Policy (loopback-only baseUrl, path approval) is
// enforced where each value is used, so a hand-edited file is not discarded
// just for naming a non-local endpoint. passthrough() keeps fields written by
// a newer version so a save does not drop them.
const providerFields = {
  kind: z.enum(['anthropic', 'openai', 'openai-compatible', 'embedded']),
  model: z.string(),
  apiKey: z.string().optional(),
  baseUrl: z.string().optional(),
  modelPath: z.string().optional(),
  contextSize: z.number().optional(),
  threads: z.number().optional(),
  gpuLayers: z.number().optional(),
};
const settingsFileSchema = z.object({
  providers: z.array(z.object({ id: z.string().min(1), ...providerFields }).passthrough()).optional(),
  activeProviderId: z.string().nullable().optional(),
  mcpServers: z
    .array(
      z
        .object({
          id: z.string().min(1),
          name: z.string(),
          enabled: z.boolean(),
          transport: z.enum(['builtin', 'stdio']),
          type: z.enum(['sqlite', 'csv-directory', 'filesystem']).optional(),
          command: z.string().optional(),
          args: z.array(z.string()).optional(),
          env: z.record(z.string()).optional(),
          allowedPaths: z.array(z.string()),
          options: z.record(z.unknown()).optional(),
        })
        .passthrough(),
    )
    .optional(),
  approvedPaths: z.array(z.string()).optional(),
  /** Legacy single-provider shape (pre model-list). */
  provider: z.object(providerFields).passthrough().optional(),
});

/** Why the last load could not use settings.json. Shown in the taskpane until dismissed. */
export interface SettingsProblem {
  kind: 'unreadable' | 'invalid';
  detail: string;
  /** Where the original file was moved; it is never deleted. */
  preservedAs: string;
  at: string;
}

/** The settings file exists but could not be read or moved aside; nothing was changed. */
export class SettingsUnavailableError extends Error {
  constructor(file: string, cause: unknown) {
    super(
      `Settings file ${file} could not be read (${cause instanceof Error ? cause.message : String(cause)}); ` +
        'refusing to continue with empty settings',
    );
    this.name = 'SettingsUnavailableError';
  }
}

let lastProblem: SettingsProblem | null = null;

export function settingsProblem(): SettingsProblem | null {
  return lastProblem;
}

export function clearSettingsProblem(): void {
  lastProblem = null;
}

export function dataDir(): string {
  const dir = process.env.EXCELAI_DATA_DIR ?? join(homedir(), '.excelai');
  mkdirSync(dir, { recursive: true });
  return dir;
}

const SETTINGS_FILE = 'settings.json';

function emptySettings(): Settings {
  return { providers: [], activeProviderId: null, mcpServers: [], approvedPaths: [] };
}

/**
 * Settings live in a 0600 JSON file in the data directory. (OS keychain
 * integration is planned; a file owned by the user's account is the honest
 * v1 — it is never sent anywhere and never committed.)
 *
 * A missing file means a fresh install. A file that does not parse, or parses
 * to the wrong shape, is moved aside to `settings.json.corrupt-<ts>` and
 * reported through {@link settingsProblem}, so no later save can overwrite
 * it. A file that exists but cannot be read (locked, permissions) throws
 * {@link SettingsUnavailableError}: starting empty would let the next save
 * replace settings that are probably fine.
 */
export function loadSettings(): Settings {
  const file = join(dataDir(), SETTINGS_FILE);
  let text: string;
  try {
    text = readFileSync(file, 'utf8');
  } catch (e) {
    if ((e as NodeJS.ErrnoException).code === 'ENOENT') return emptySettings();
    throw new SettingsUnavailableError(file, e);
  }
  let json: unknown;
  try {
    json = JSON.parse(text);
  } catch (e) {
    quarantine(file, 'unreadable', `not valid JSON: ${e instanceof Error ? e.message : String(e)}`);
    return emptySettings();
  }
  const parsed = settingsFileSchema.safeParse(json);
  if (!parsed.success) {
    const detail = parsed.error.issues
      .slice(0, 3)
      .map((i) => `${i.path.join('.') || '(root)'}: ${i.message}`)
      .join('; ');
    quarantine(file, 'invalid', detail);
    return emptySettings();
  }
  const raw = parsed.data as {
    providers?: StoredProvider[];
    activeProviderId?: string | null;
    mcpServers?: McpServerConfig[];
    approvedPaths?: string[];
    provider?: ProviderConfig;
  };
  // Migrate the legacy single-provider shape.
  if (!raw.providers && raw.provider) {
    const migrated: StoredProvider = { ...raw.provider, id: randomUUID() };
    const settings: Settings = {
      providers: [migrated],
      activeProviderId: migrated.id,
      mcpServers: raw.mcpServers ?? [],
      approvedPaths: raw.approvedPaths ?? [],
    };
    saveSettings(settings);
    return settings;
  }
  const providers = raw.providers ?? [];
  const activeProviderId =
    raw.activeProviderId && providers.some((p) => p.id === raw.activeProviderId)
      ? raw.activeProviderId
      : (providers[0]?.id ?? null);
  const mcpServers = raw.mcpServers ?? [];
  const approvedPaths = raw.approvedPaths ?? [];
  return { providers, activeProviderId, mcpServers, approvedPaths };
}

function quarantine(file: string, kind: SettingsProblem['kind'], detail: string): void {
  const at = new Date().toISOString();
  const preservedAs = `${file}.corrupt-${at.replace(/[:.]/g, '-')}`;
  try {
    renameSync(file, preservedAs);
  } catch (e) {
    // Could not move it aside: leave it in place and refuse, rather than
    // return empty settings that the next save would write over it.
    throw new SettingsUnavailableError(file, e);
  }
  lastProblem = { kind, detail, preservedAs, at };
  console.warn(`[settings] ${file} is ${kind} (${detail}); moved to ${preservedAs}, starting with empty settings`);
}

/**
 * Atomic replace: write a temp file next to the target, fsync it, then rename
 * it over the target. A crash leaves the old file or the new one, never a
 * truncated mix.
 */
export function saveSettings(settings: Settings): void {
  const file = join(dataDir(), SETTINGS_FILE);
  const tmp = `${file}.${process.pid}.${randomUUID()}.tmp`;
  const fd = openSync(tmp, 'wx', 0o600);
  try {
    writeFileSync(fd, JSON.stringify(settings, null, 2));
    fsyncSync(fd);
  } finally {
    closeSync(fd);
  }
  try {
    renameSync(tmp, file);
  } catch (e) {
    rmSync(tmp, { force: true });
    throw e;
  }
  // Persist the rename itself. Windows cannot open a directory to fsync it.
  if (process.platform !== 'win32') {
    const dirFd = openSync(dirname(file), 'r');
    try {
      fsyncSync(dirFd);
    } finally {
      closeSync(dirFd);
    }
  }
}

/**
 * Stable, non-secret identity of a provider configuration, used to namespace
 * cached function results: two saved configs never share entries, and editing
 * the endpoint, model or model file starts a fresh namespace. The API key is
 * left out on purpose: it is a credential, not part of what produces answers.
 */
export function providerIdentity(p: StoredProvider): string {
  const { apiKey: _credential, ...rest } = p;
  const canonical = Object.fromEntries(Object.entries(rest).sort(([x], [y]) => x.localeCompare(y)));
  return createHash('sha256').update(JSON.stringify(canonical)).digest('hex');
}

export function activeProvider(settings: Settings): StoredProvider | null {
  return settings.providers.find((p) => p.id === settings.activeProviderId) ?? null;
}

export interface RedactedProvider {
  id: string;
  kind: StoredProvider['kind'];
  model: string;
  baseUrl?: string;
  /** Masked: present only to show that a key is stored. */
  apiKey?: string;
  modelPath?: string;
}

export function redactProvider(p: StoredProvider): RedactedProvider {
  const out: RedactedProvider = { id: p.id, kind: p.kind, model: p.model };
  if (p.baseUrl) out.baseUrl = p.baseUrl;
  if (p.apiKey) out.apiKey = `••••${p.apiKey.slice(-4)}`;
  if (p.modelPath) out.modelPath = p.modelPath;
  return out;
}
