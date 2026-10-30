import { createServer, type IncomingMessage, type Server, type ServerResponse } from 'node:http';
import { createServer as createHttpsServer } from 'node:https';
import { existsSync, readFileSync, statSync } from 'node:fs';
import { extname, join, normalize } from 'node:path';
import { randomUUID } from 'node:crypto';
import { WebSocketServer, type WebSocket } from 'ws';
import { PROTOCOL_VERSION, clientMessageSchema, makeEnvelope, type McpServerConfig } from '@excelai/shared';
import { z } from 'zod';
import { Auth } from './auth.js';
import {
  activeProvider,
  loadSettings,
  redactProvider,
  saveSettings,
  type StoredProvider,
} from './config.js';
import type { Db } from './db.js';
import { Ledger } from './ledger.js';
import { MemoryStore } from './memory/store.js';
import { FunctionEngine } from './functions/batch.js';
import { createProvider, ProviderError, getModelsDir, listLocalGgufModels, type Provider } from './providers/index.js';
import { BridgeExecutor } from './bridge.js';
import { Session } from './session.js';
import { assertLoopbackUrl } from './networkGuard.js';
import { loadAndParseLocalFile, parseTabularContent } from './importers/tabular.js';
import { McpManager, PathGuard } from './mcp/index.js';

const SIDECAR_VERSION = '0.1.0';
const MAX_CACHED_SESSIONS = 50;
// Largest legitimate frame is a tool_result carrying a range snapshot; 16 MiB
// leaves headroom over MAX_SNAPSHOT_CELLS while staying far below ws's 100 MiB default.
const MAX_WS_PAYLOAD_BYTES = 16 * 1024 * 1024;
/** Malformed frames tolerated per socket; the next one closes it (1008). */
const MAX_PROTOCOL_VIOLATIONS = 3;

/** Origins allowed to talk to the sidecar (Office webviews + local dev). */
const ALLOWED_ORIGIN_RE =
  /^(https:\/\/[a-z0-9-]+\.officeapps\.live\.com|https:\/\/[a-z0-9-]+\.office\.com|https:\/\/[a-z0-9.-]+\.sharepoint\.com|https?:\/\/localhost(:\d+)?|https?:\/\/127\.0\.0\.1(:\d+)?)$/i;

export interface SidecarServer {
  server: Server;
  auth: Auth;
  port: number;
  close(): Promise<void>;
}

const providerConfigSchema = z.object({
  kind: z.enum(['anthropic', 'openai', 'openai-compatible', 'embedded']),
  model: z.string().min(1),
  apiKey: z.string().min(1).optional(),
  baseUrl: z
    .string()
    .url()
    .optional()
    .refine(
      (u) => {
        if (!u) return true;
        try {
          assertLoopbackUrl(u);
          return true;
        } catch {
          return false;
        }
      },
      { message: 'baseUrl must be a local loopback address (127.0.0.1, localhost, [::1])' },
    ),
  modelPath: z.string().optional(),
  contextSize: z.number().int().positive().optional(),
  threads: z.number().int().positive().optional(),
  gpuLayers: z.number().int().min(0).optional(),
});

const mcpServerConfigSchema = z.object({
  id: z.string().optional(),
  name: z.string().min(1),
  enabled: z.boolean().default(true),
  transport: z.enum(['builtin', 'stdio']),
  type: z.enum(['sqlite', 'csv-directory', 'filesystem']).optional(),
  command: z.string().optional(),
  args: z.array(z.string()).optional(),
  env: z.record(z.string()).optional(),
  allowedPaths: z.array(z.string()).default([]),
  options: z.record(z.unknown()).optional(),
});

export async function startServer(options: {
  db: Db;
  port?: number;
  staticDir?: string;
}): Promise<SidecarServer> {
  const { db } = options;
  const auth = new Auth(db);
  const ledger = new Ledger(db);
  const memory = new MemoryStore(db);
  const functions = new FunctionEngine(db);
  const mcpManager = new McpManager();

  const initialSettings = loadSettings();
  await mcpManager.sync(initialSettings.mcpServers, initialSettings.approvedPaths);

  let cachedProvider: { provider: Provider; fingerprint: string } | null = null;
  const getProvider = (): Provider => {
    const active = activeProvider(loadSettings()) ?? {
      id: 'default-lm-studio',
      kind: 'openai-compatible',
      model: 'local-model',
      baseUrl: 'http://127.0.0.1:1234/v1',
      apiKey: 'lm-studio',
    };
    const fingerprint = JSON.stringify(active);
    if (!cachedProvider || cachedProvider.fingerprint !== fingerprint) {
      cachedProvider = { provider: createProvider(active), fingerprint };
    }
    return cachedProvider.provider;
  };

  // Optional TLS (e.g. office-addin-dev-certs for hosts that require
  // https://localhost): set EXCELAI_TLS_CERT and EXCELAI_TLS_KEY.
  const tlsCert = process.env.EXCELAI_TLS_CERT;
  const tlsKey = process.env.EXCELAI_TLS_KEY;
  const handler = (req: IncomingMessage, res: ServerResponse): void => {
    void handleHttp(req, res).catch((e) => {
      sendJson(res, 500, { error: e instanceof Error ? e.message : 'Internal error' });
    });
  };
  const server: Server =
    tlsCert && tlsKey
      ? createHttpsServer({ cert: readFileSync(tlsCert), key: readFileSync(tlsKey) }, handler)
      : createServer(handler);

  async function handleHttp(req: IncomingMessage, res: ServerResponse): Promise<void> {
    const url = new URL(req.url ?? '/', 'http://127.0.0.1');
    const origin = req.headers.origin;

    // CORS for the Office webview / browser clients.
    if (origin && ALLOWED_ORIGIN_RE.test(origin)) {
      res.setHeader('access-control-allow-origin', origin);
      res.setHeader('access-control-allow-headers', 'authorization, content-type');
      res.setHeader('access-control-allow-methods', 'GET, POST, PUT, DELETE, OPTIONS');
    }
    if (req.method === 'OPTIONS') {
      res.writeHead(204).end();
      return;
    }

    if (url.pathname === '/api/health') {
      sendJson(res, 200, { ok: true, version: SIDECAR_VERSION, protocol: PROTOCOL_VERSION });
      return;
    }

    if (url.pathname === '/api/pair' && req.method === 'POST') {
      const body = await readJson(req);
      const code = typeof (body as { code?: unknown })?.code === 'string' ? (body as { code: string }).code : '';
      const token = auth.pair(code);
      if (!token) {
        sendJson(res, 403, { error: 'Invalid or already-used pairing code' });
        return;
      }
      sendJson(res, 200, { token });
      return;
    }

    if (url.pathname.startsWith('/api/')) {
      const token = bearerToken(req);
      if (!auth.verify(token)) {
        sendJson(res, 401, { error: 'Unauthorized' });
        return;
      }

      if (url.pathname === '/api/auth/check' && req.method === 'GET') {
        sendJson(res, 200, { ok: true });
        return;
      }

      /* ----- saved model configurations ----- */

      if (url.pathname === '/api/settings/providers' && req.method === 'GET') {
        const settings = loadSettings();
        sendJson(res, 200, {
          providers: settings.providers.map(redactProvider),
          activeProviderId: settings.activeProviderId,
        });
        return;
      }
      if (url.pathname === '/api/settings/providers' && req.method === 'POST') {
        const parsed = providerConfigSchema.safeParse(await readJson(req));
        if (!parsed.success) {
          sendJson(res, 400, { error: parsed.error.issues.map((i) => i.message).join('; ') });
          return;
        }
        const settings = loadSettings();
        const stored: StoredProvider = { ...parsed.data, id: crypto.randomUUID() };
        settings.providers.push(stored);
        // First saved model becomes active automatically.
        if (!settings.activeProviderId) settings.activeProviderId = stored.id;
        saveSettings(settings);
        cachedProvider = null;
        sendJson(res, 200, { id: stored.id, activeProviderId: settings.activeProviderId });
        return;
      }

      const providerRoute = /^\/api\/settings\/providers\/([\w-]+)(?:\/(activate|test))?$/.exec(
        url.pathname,
      );
      if (providerRoute) {
        const [, id, action] = providerRoute;
        const settings = loadSettings();
        const stored = settings.providers.find((p) => p.id === id);

        if (!action && req.method === 'DELETE') {
          if (!stored) {
            sendJson(res, 404, { error: 'No such model configuration' });
            return;
          }
          settings.providers = settings.providers.filter((p) => p.id !== id);
          if (settings.activeProviderId === id) {
            settings.activeProviderId = settings.providers[0]?.id ?? null;
          }
          saveSettings(settings);
          cachedProvider = null;
          sendJson(res, 200, { ok: true, activeProviderId: settings.activeProviderId });
          return;
        }
        if (action === 'activate' && req.method === 'POST') {
          if (!stored) {
            sendJson(res, 404, { error: 'No such model configuration' });
            return;
          }
          settings.activeProviderId = stored.id;
          saveSettings(settings);
          cachedProvider = null;
          sendJson(res, 200, { ok: true, activeProviderId: stored.id });
          return;
        }
        if (action === 'test' && req.method === 'POST') {
          if (!stored) {
            sendJson(res, 404, { ok: false, error: 'No such model configuration' });
            return;
          }
          sendJson(res, 200, await testProviderConfig(stored, ledger));
          return;
        }
      }

      /** Test a candidate configuration before saving it. */
      if (url.pathname === '/api/settings/providers/test' && req.method === 'POST') {
        const parsed = providerConfigSchema.safeParse(await readJson(req));
        if (!parsed.success) {
          sendJson(res, 400, { ok: false, error: parsed.error.issues.map((i) => i.message).join('; ') });
          return;
        }
        sendJson(res, 200, await testProviderConfig(parsed.data, ledger));
        return;
      }

      /** Best-effort model discovery for the add/edit form (and an Ollama
       *  readiness check: lists the models actually installed). */
      if (url.pathname === '/api/providers/models' && req.method === 'POST') {
        const schema = z.object({
          kind: z.enum(['anthropic', 'openai', 'openai-compatible', 'embedded']),
          apiKey: z.string().optional(),
          baseUrl: z
            .string()
            .url()
            .optional()
            .refine(
              (u) => {
                if (!u) return true;
                try {
                  assertLoopbackUrl(u);
                  return true;
                } catch {
                  return false;
                }
              },
              { message: 'baseUrl must be a local loopback address' },
            ),
        });
        const parsed = schema.safeParse(await readJson(req));
        if (!parsed.success) {
          sendJson(res, 400, { models: [], error: parsed.error.issues.map((i) => i.message).join('; ') });
          return;
        }
        sendJson(res, 200, await listModels(parsed.data));
        return;
      }
      if (url.pathname === '/api/functions/run' && req.method === 'POST') {
        const schema = z.object({
          kind: z.enum(['ai', 'classify', 'extract', 'translate']),
          prompt: z.string().default(''),
          input: z.string().optional(),
          categories: z.array(z.string()).optional(),
          targetLang: z.string().optional(),
        });
        const parsed = schema.safeParse(await readJson(req));
        if (!parsed.success) {
          sendJson(res, 400, { error: 'Invalid function request' });
          return;
        }
        const controller = new AbortController();
        req.on('close', () => {
          if (!res.writableEnded) {
            controller.abort();
          }
        });
        try {
          const provider = getProvider();
          const value = await functions.run(
            provider,
            parsed.data,
            (usage) =>
              ledger.record({ provider: provider.id, model: provider.model, feature: 'functions', ...usage }),
            controller.signal,
          );
          sendJson(res, 200, { value });
        } catch (e) {
          if (
            controller.signal.aborted ||
            (e instanceof Error && (e.name === 'AbortError' || e.message.toLowerCase().includes('abort')))
          ) {
            if (!res.writableEnded) {
              sendJson(res, 499, { error: 'Calculation canceled' });
            }
          } else {
            sendJson(res, 502, { error: e instanceof Error ? e.message : 'Function failed' });
          }
        }
        return;
      }
      if (url.pathname === '/api/importer/parse' && req.method === 'POST') {
        const schema = z
          .object({
            filePath: z.string().optional(),
            content: z.string().optional(),
            format: z.enum(['csv', 'tsv', 'markdown', 'json', 'auto']).optional(),
          })
          .refine((d) => d.filePath || d.content, {
            message: 'Either filePath or content must be provided',
          });
        const parsed = schema.safeParse(await readJson(req));
        if (!parsed.success) {
          sendJson(res, 400, { ok: false, error: parsed.error.issues.map((i) => i.message).join('; ') });
          return;
        }
        try {
          if (parsed.data.filePath) {
            const result = loadAndParseLocalFile(parsed.data.filePath);
            const headers = result.rows[0]?.map((c) => String(c ?? '')) ?? [];
            sendJson(res, 200, {
              ok: true,
              format: result.format,
              totalRows: result.totalRows,
              totalCols: result.totalCols,
              headers,
              previewRows: result.rows.slice(0, 10),
            });
          } else if (parsed.data.content) {
            const rows = parseTabularContent(parsed.data.content, parsed.data.format);
            const headers = rows[0]?.map((c) => String(c ?? '')) ?? [];
            sendJson(res, 200, {
              ok: true,
              format: parsed.data.format ?? 'auto',
              totalRows: rows.length,
              totalCols: rows[0]?.length ?? 0,
              headers,
              previewRows: rows.slice(0, 10),
            });
          }
        } catch (e) {
          sendJson(res, 400, { ok: false, error: e instanceof Error ? e.message : String(e) });
        }
        return;
      }
      if (url.pathname === '/api/ledger/summary' && req.method === 'GET') {
        sendJson(res, 200, ledger.summary());
        return;
      }
      if (url.pathname === '/api/memory' && req.method === 'GET') {
        const scope = url.searchParams.get('scope') ?? undefined;
        sendJson(res, 200, { entries: memory.list(scope) });
        return;
      }
      if (url.pathname === '/api/memory' && req.method === 'POST') {
        const schema = z.object({ scope: z.string().min(1), content: z.string().min(1) });
        const parsed = schema.safeParse(await readJson(req));
        if (!parsed.success) {
          sendJson(res, 400, { error: 'scope and content are required' });
          return;
        }
        sendJson(res, 200, memory.add(parsed.data.scope, parsed.data.content));
        return;
      }

      /* ----- Model Context Protocol (MCP) ----- */

      if (url.pathname === '/api/mcp/servers' && req.method === 'GET') {
        const settings = loadSettings();
        sendJson(res, 200, {
          servers: mcpManager.getServerStatuses(settings.mcpServers),
        });
        return;
      }

      if (url.pathname === '/api/mcp/servers' && req.method === 'POST') {
        const parsed = mcpServerConfigSchema.safeParse(await readJson(req));
        if (!parsed.success) {
          sendJson(res, 400, { ok: false, error: parsed.error.issues.map((i) => i.message).join('; ') });
          return;
        }
        const settings = loadSettings();
        const serverData: McpServerConfig = {
          ...parsed.data,
          id: parsed.data.id || randomUUID(),
        };
        const idx = settings.mcpServers.findIndex((s) => s.id === serverData.id);
        if (idx >= 0) {
          settings.mcpServers[idx] = serverData;
        } else {
          settings.mcpServers.push(serverData);
        }
        saveSettings(settings);
        await mcpManager.sync(settings.mcpServers, settings.approvedPaths);
        sendJson(res, 200, { ok: true, server: serverData });
        return;
      }

      const mcpServerToggleMatch = url.pathname.match(/^\/api\/mcp\/servers\/([^/]+)\/toggle$/);
      if (mcpServerToggleMatch && req.method === 'POST') {
        const serverId = mcpServerToggleMatch[1];
        const settings = loadSettings();
        const server = settings.mcpServers.find((s) => s.id === serverId);
        if (!server) {
          sendJson(res, 404, { ok: false, error: 'MCP Server not found' });
          return;
        }
        server.enabled = !server.enabled;
        saveSettings(settings);
        await mcpManager.sync(settings.mcpServers, settings.approvedPaths);
        sendJson(res, 200, { ok: true, enabled: server.enabled });
        return;
      }

      const mcpServerDeleteMatch = url.pathname.match(/^\/api\/mcp\/servers\/([^/]+)$/);
      if (mcpServerDeleteMatch && req.method === 'DELETE') {
        const serverId = mcpServerDeleteMatch[1];
        const settings = loadSettings();
        settings.mcpServers = settings.mcpServers.filter((s) => s.id !== serverId);
        saveSettings(settings);
        await mcpManager.sync(settings.mcpServers, settings.approvedPaths);
        sendJson(res, 200, { ok: true });
        return;
      }

      if (url.pathname === '/api/mcp/servers/test' && req.method === 'POST') {
        const parsed = mcpServerConfigSchema.safeParse(await readJson(req));
        if (!parsed.success) {
          sendJson(res, 400, { ok: false, error: parsed.error.issues.map((i) => i.message).join('; ') });
          return;
        }
        const settings = loadSettings();
        const testRes = await mcpManager.testServer(
          { ...parsed.data, id: parsed.data.id || 'test-temp' },
          settings.approvedPaths,
        );
        sendJson(res, 200, testRes);
        return;
      }

      if (url.pathname === '/api/mcp/paths' && req.method === 'GET') {
        const settings = loadSettings();
        sendJson(res, 200, { paths: settings.approvedPaths });
        return;
      }

      if (url.pathname === '/api/mcp/paths' && req.method === 'POST') {
        const body = (await readJson(req)) as { path?: string };
        const rawPath = typeof body?.path === 'string' ? body.path.trim() : '';
        if (!rawPath) {
          sendJson(res, 400, { ok: false, error: 'Path is required' });
          return;
        }
        let canonical: string;
        try {
          canonical = PathGuard.canonicalize(rawPath);
        } catch (e) {
          sendJson(res, 400, { ok: false, error: e instanceof Error ? e.message : String(e) });
          return;
        }
        const settings = loadSettings();
        if (!settings.approvedPaths.includes(canonical)) {
          settings.approvedPaths.push(canonical);
          saveSettings(settings);
          await mcpManager.sync(settings.mcpServers, settings.approvedPaths);
        }
        sendJson(res, 200, { ok: true, paths: settings.approvedPaths });
        return;
      }

      if (url.pathname === '/api/mcp/paths' && req.method === 'DELETE') {
        const body = (await readJson(req)) as { path?: string };
        const rawPath = typeof body?.path === 'string' ? body.path.trim() : '';
        const settings = loadSettings();
        settings.approvedPaths = settings.approvedPaths.filter((p) => p !== rawPath);
        saveSettings(settings);
        await mcpManager.sync(settings.mcpServers, settings.approvedPaths);
        sendJson(res, 200, { ok: true, paths: settings.approvedPaths });
        return;
      }

      if (url.pathname === '/api/mcp/tools' && req.method === 'GET') {
        sendJson(res, 200, { tools: mcpManager.getToolSpecs() });
        return;
      }
      sendJson(res, 404, { error: 'Not found' });
      return;
    }

    // Static add-in assets (same-origin serving avoids mixed-content issues).
    if (options.staticDir) {
      serveStatic(options.staticDir, url.pathname, res);
      return;
    }
    sendJson(res, 404, { error: 'Not found' });
  }

  /* ---------------- WebSocket ---------------- */

  const wss = new WebSocketServer({ noServer: true, maxPayload: MAX_WS_PAYLOAD_BYTES });
  const sessions = new Map<string, Session>();

  server.on('upgrade', (req, socket, head) => {
    const url = new URL(req.url ?? '/', 'http://127.0.0.1');
    const origin = req.headers.origin;
    if (url.pathname !== '/ws' || (origin && !ALLOWED_ORIGIN_RE.test(origin))) {
      socket.destroy();
      return;
    }
    if (!auth.verify(url.searchParams.get('token'))) {
      socket.destroy();
      return;
    }
    wss.handleUpgrade(req, socket, head, (ws) => wss.emit('connection', ws, req));
  });

  wss.on('connection', (ws: WebSocket) => {
    let session: Session | null = null;
    let bridge: BridgeExecutor | null = null;
    let violations = 0;

    ws.on('error', (err) => {
      console.error('[ws] socket error:', err.message);
    });

    /** Answer a malformed frame with a typed error; repeated offenders are closed. */
    const rejectFrame = (message: string, issues: string[], refId?: string): void => {
      violations += 1;
      ws.send(
        JSON.stringify(makeEnvelope('protocol_error', { message, issues: issues.slice(0, 10), refId })),
      );
      if (violations >= MAX_PROTOCOL_VIOLATIONS) ws.close(1008, 'Too many malformed frames');
    };

    ws.on('message', (raw, isBinary) => {
      (async () => {
        if (isBinary) {
          rejectFrame('Binary frames are not supported', []);
          return;
        }
        let parsed: unknown;
        try {
          parsed = JSON.parse(String(raw));
        } catch {
          rejectFrame('Frame is not valid JSON', []);
          return;
        }
        const frame = (typeof parsed === 'object' && parsed !== null ? parsed : {}) as Record<string, unknown>;
        const refId = typeof frame.id === 'string' ? frame.id.slice(0, 128) : undefined;
        if (typeof frame.v === 'number' && Math.floor(frame.v) !== PROTOCOL_VERSION) {
          ws.send(
            JSON.stringify({
              v: PROTOCOL_VERSION,
              id: refId ?? 'unknown',
              type: 'hello_ack',
              payload: {
                ok: false,
                error: `Protocol version mismatch (sidecar ${PROTOCOL_VERSION}, client ${String(frame.v)}). Update both sides.`,
                sidecarVersion: SIDECAR_VERSION,
              },
            }),
          );
          ws.close();
          return;
        }
        const result = clientMessageSchema.safeParse(parsed);
        if (!result.success) {
          rejectFrame(
            'Malformed frame',
            result.error.issues.map((i) => `${i.path.join('.') || '(frame)'}: ${i.message}`),
            refId,
          );
          return;
        }
        const msg = result.data;
        switch (msg.type) {
          case 'hello': {
            if (!auth.verify(msg.payload.token)) {
              ws.close();
              return;
            }
            bridge = new BridgeExecutor(ws);
            const existingSession = sessions.get(msg.payload.workbookId);
            if (existingSession) {
              existingSession.rebind(ws, msg.payload.workbookName, bridge);
              session = existingSession;
              // Refresh insertion order for LRU bounding
              sessions.delete(msg.payload.workbookId);
              sessions.set(msg.payload.workbookId, existingSession);
            } else {
              session = new Session(
                ws,
                msg.payload.workbookId,
                msg.payload.workbookName,
                bridge,
                getProvider,
                ledger,
                memory,
                mcpManager,
              );
              if (sessions.size >= MAX_CACHED_SESSIONS) {
                const oldestKey = sessions.keys().next().value;
                if (oldestKey !== undefined) sessions.delete(oldestKey);
              }
              sessions.set(msg.payload.workbookId, session);
            }
            ws.send(
              JSON.stringify({
                v: PROTOCOL_VERSION,
                id: msg.id,
                type: 'hello_ack',
                payload: { ok: true, sidecarVersion: SIDECAR_VERSION },
              }),
            );
            break;
          }
          case 'tool_result':
            bridge?.handleToolResult(msg.payload);
            break;
          case 'chat':
            await session?.handleChat(msg.payload.message);
            break;
          case 'abort':
            session?.handleAbort(msg.payload.reason);
            break;
          case 'approve_changeset':
            await session?.handleApprove(msg.payload.changeSetId);
            break;
          case 'reject_changeset':
            session?.handleReject(msg.payload.changeSetId);
            break;
          case 'undo_changeset':
            await session?.handleUndo(msg.payload.changeSetId);
            break;
        }
      })().catch((err: unknown) => {
        // A bad frame from one client must never take down the process.
        console.error('[ws] message handler failed:', err instanceof Error ? err.message : err);
        ws.close(1011, 'Internal error');
      });
    });

    ws.on('close', () => {
      if (session?.isCurrentSocket(ws)) {
        session.handleAbort('Add-in disconnected');
      }
      bridge?.abortAll('Add-in disconnected');
    });
  });

  const port = options.port ?? 8923;
  await new Promise<void>((resolve, reject) => {
    server.once('error', reject);
    server.listen(port, '127.0.0.1', resolve);
  });
  const actualPort = (server.address() as { port: number }).port;

  return {
    server,
    auth,
    port: actualPort,
    close: async () => {
      await mcpManager.closeAll();
      wss.close();
      await new Promise<void>((resolve, reject) => {
        server.close((e) => (e ? reject(e) : resolve()));
      });
    },
  };
}

/* ---------------- helpers ---------------- */

async function testProviderConfig(
  config: Parameters<typeof createProvider>[0],
  ledger: Ledger,
): Promise<{ ok: boolean; reply?: string; error?: string }> {
  try {
    const provider = createProvider(config);
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 20_000);
    const response = await provider
      .chat({
        messages: [{ role: 'user', content: 'Reply with exactly: OK' }],
        maxTokens: 16,
        temperature: 0,
        abortSignal: controller.signal,
      })
      .finally(() => clearTimeout(timer));
    ledger.record({
      provider: provider.id,
      model: provider.model,
      feature: 'probe',
      ...response.usage,
    });
    return { ok: true, reply: response.text.trim().slice(0, 80) };
  } catch (e) {
    const message =
      e instanceof Error && e.name === 'AbortError'
        ? 'Timed out after 20s — is the endpoint reachable?'
        : e instanceof Error
          ? e.message
          : String(e);
    return { ok: false, error: message };
  }
}

async function listModels(config: {
  kind: 'anthropic' | 'openai' | 'openai-compatible' | 'embedded';
  apiKey?: string;
  baseUrl?: string;
}): Promise<{ models: string[]; error?: string }> {
  if (config.kind === 'embedded') {
    const models = await listLocalGgufModels();
    if (models.length === 0) {
      return {
        models: [],
        error: `No .gguf models found in "${getModelsDir()}". Place GGUF model files there to run in-process.`,
      };
    }
    return { models };
  }
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 8_000);
  try {
    const rawBase = config.baseUrl || 'http://127.0.0.1:1234/v1';
    assertLoopbackUrl(rawBase);
    let endpoint: string;
    const headers: Record<string, string> = {};
    if (config.kind === 'anthropic') {
      endpoint = `${rawBase.replace(/\/$/, '')}/v1/models`;
      if (config.apiKey) headers['x-api-key'] = config.apiKey;
      headers['anthropic-version'] = '2023-06-01';
    } else {
      endpoint = `${rawBase.replace(/\/$/, '')}/models`;
      const key = config.apiKey || 'lm-studio';
      headers.authorization = `Bearer ${key}`;
    }
    assertLoopbackUrl(endpoint);
    // A redirect could point anywhere; the loopback check above only covers the first hop.
    const res = await fetch(endpoint, { headers, signal: controller.signal, redirect: 'error' });
    if (!res.ok) return { models: [], error: `Endpoint answered ${res.status}` };
    const data = (await res.json()) as { data?: { id?: string }[]; models?: { id?: string }[] };
    const items = data.data ?? data.models ?? [];
    const models = items
      .map((m) => m.id)
      .filter((id): id is string => typeof id === 'string')
      .sort();
    return { models };
  } catch (e) {
    return {
      models: [],
      error:
        e instanceof Error && e.name === 'AbortError'
          ? 'Timed out — endpoint unreachable'
          : e instanceof Error
            ? e.message
            : String(e),
    };
  } finally {
    clearTimeout(timer);
  }
}

function bearerToken(req: IncomingMessage): string | null {
  const header = req.headers.authorization;
  if (header?.startsWith('Bearer ')) return header.slice(7);
  return null;
}

function sendJson(res: ServerResponse, status: number, body: unknown): void {
  const json = JSON.stringify(body);
  res.writeHead(status, { 'content-type': 'application/json' });
  res.end(json);
}

async function readJson(req: IncomingMessage): Promise<unknown> {
  const chunks: Buffer[] = [];
  let size = 0;
  for await (const chunk of req) {
    size += (chunk as Buffer).length;
    if (size > 5_000_000) throw new Error('Request body too large');
    chunks.push(chunk as Buffer);
  }
  const text = Buffer.concat(chunks).toString('utf8');
  if (!text) return {};
  try {
    return JSON.parse(text);
  } catch {
    return {};
  }
}

const MIME: Record<string, string> = {
  '.html': 'text/html',
  '.js': 'text/javascript',
  '.css': 'text/css',
  '.json': 'application/json',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
};

function serveStatic(dir: string, pathname: string, res: ServerResponse): void {
  const rel = pathname === '/' ? '/index.html' : pathname;
  const file = normalize(join(dir, rel));
  if (!file.startsWith(normalize(dir)) || !existsSync(file) || !statSync(file).isFile()) {
    sendJson(res, 404, { error: 'Not found' });
    return;
  }
  res.writeHead(200, { 'content-type': MIME[extname(file)] ?? 'application/octet-stream' });
  res.end(readFileSync(file));
}
