import { mkdtempSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { dataDir } from '../src/config.js';
import { openDb, pruneFnCache, type RetentionLimits } from '../src/db.js';
import { Ledger } from '../src/ledger.js';
import { FunctionEngine, FunctionQueueFullError } from '../src/functions/batch.js';
import { ChangeSetManager, DEFAULT_CHANGESET_RETENTION } from '../src/changeset/manager.js';
import { ScriptedProvider } from '../src/providers/scripted.js';
import { WorkbookEmulator } from '../src/workbook/emulator.js';
import type { ChangeOp } from '@excelai/shared';

const DAY = 24 * 60 * 60 * 1000;
const limits = (over: Partial<RetentionLimits>): RetentionLimits => ({
  fnCacheMaxRows: 1_000,
  fnCacheMaxAgeMs: 30 * DAY,
  ledgerMaxRows: 1_000,
  ledgerMaxAgeMs: 400 * DAY,
  ...over,
});

afterEach(() => {
  vi.useRealTimers();
});

const count = (db: ReturnType<typeof openDb>, table: string): number =>
  (db.prepare(`SELECT COUNT(*) AS n FROM ${table}`).get() as { n: number }).n;

describe('function cache retention (L3)', () => {
  it('keeps at most fnCacheMaxRows entries, dropping the oldest', async () => {
    const db = openDb(':memory:');
    const engine = new FunctionEngine(db, { retention: limits({ fnCacheMaxRows: 3 }) });
    const provider = new ScriptedProvider(Array.from({ length: 6 }, (_, i) => ({ text: `v${i}` })));
    for (let i = 0; i < 5; i++) {
      vi.useFakeTimers({ toFake: ['Date'] });
      vi.setSystemTime(new Date(Date.UTC(2026, 0, 1, 0, 0, i)));
      await engine.run(provider, 'cfg', { kind: 'ai', prompt: `p${i}` });
      vi.useRealTimers();
    }
    expect(count(db, 'fn_cache')).toBe(3);
    // p0 was evicted, so it is computed again; p4 is still cached.
    expect(await engine.run(provider, 'cfg', { kind: 'ai', prompt: 'p4' })).toBe('v4');
    expect(await engine.run(provider, 'cfg', { kind: 'ai', prompt: 'p0' })).toBe('v5');
  });

  it('expires entries older than fnCacheMaxAgeMs', () => {
    const db = openDb(':memory:');
    const insert = db.prepare('INSERT INTO fn_cache (key, value, created_at) VALUES (?, ?, ?)');
    insert.run('old', 'x', new Date(Date.now() - 31 * DAY).toISOString());
    insert.run('new', 'y', new Date().toISOString());
    pruneFnCache(db);
    expect((db.prepare('SELECT key FROM fn_cache').all() as { key: string }[]).map((r) => r.key)).toEqual(['new']);
  });

  it('prunes an existing database file when it is opened', () => {
    const path = join(mkdtempSync(join(dataDir(), 'db-')), 'excelai.db');
    const first = openDb(path);
    first
      .prepare('INSERT INTO fn_cache (key, value, created_at) VALUES (?, ?, ?)')
      .run('stale', 'x', new Date(Date.now() - 365 * DAY).toISOString());
    first
      .prepare(
        'INSERT INTO ledger (at, provider, model, feature, input_tokens, output_tokens) VALUES (?, ?, ?, ?, ?, ?)',
      )
      .run(new Date(Date.now() - 500 * DAY).toISOString(), 'p', 'm', 'chat', 1, 1);
    first.close();

    const reopened = openDb(path);
    expect(count(reopened, 'fn_cache')).toBe(0);
    expect(count(reopened, 'ledger')).toBe(0);
    reopened.close();
  });
});

describe('ledger retention (L3)', () => {
  it('keeps at most ledgerMaxRows calls and drops ones older than ledgerMaxAgeMs', () => {
    const db = openDb(':memory:');
    const ledger = new Ledger(db, limits({ ledgerMaxRows: 2 }));
    const entry = { provider: 'p', model: 'm', feature: 'chat' as const, inputTokens: 1, outputTokens: 1 };
    db.prepare(
      'INSERT INTO ledger (at, provider, model, feature, input_tokens, output_tokens) VALUES (?, ?, ?, ?, ?, ?)',
    ).run(new Date(Date.now() - 401 * DAY).toISOString(), 'p', 'm', 'chat', 100, 100);
    ledger.record(entry);
    expect(ledger.summary().totalCalls).toBe(1); // the 401-day-old row is gone
    ledger.record(entry);
    ledger.record({ ...entry, inputTokens: 5 });
    const s = ledger.summary();
    expect(s.totalCalls).toBe(2);
    expect(s.inputTokens).toBe(6); // newest two kept
  });
});

describe('function wait queue bounds (L3)', () => {
  function blockingProvider(): { provider: ScriptedProvider; release: () => void } {
    const pending: (() => void)[] = [];
    const provider = new (class extends ScriptedProvider {
      constructor() {
        super([]);
      }
      override chat(): Promise<never> {
        return new Promise((_, reject) => pending.push(() => reject(new Error('released'))));
      }
    })();
    return { provider, release: () => pending.splice(0).forEach((r) => r()) };
  }

  it('refuses calls beyond maxQueued without reaching the model', async () => {
    const engine = new FunctionEngine(openDb(':memory:'), { concurrency: 1, maxQueued: 1 });
    const { provider, release } = blockingProvider();
    const running = engine.run(provider, 'cfg', { kind: 'ai', prompt: 'a' }).catch(() => undefined);
    const queued = engine.run(provider, 'cfg', { kind: 'ai', prompt: 'b' }).catch(() => undefined);
    await expect(engine.run(provider, 'cfg', { kind: 'ai', prompt: 'c' })).rejects.toBeInstanceOf(
      FunctionQueueFullError,
    );
    expect(provider.requests).toHaveLength(0); // blocking chat() never records; nothing extra was sent
    release();
    await running;
    release();
    await queued;
  });

  it('refuses a call that waits longer than maxQueueWaitMs', async () => {
    const engine = new FunctionEngine(openDb(':memory:'), { concurrency: 1, maxQueueWaitMs: 30 });
    const { provider, release } = blockingProvider();
    const running = engine.run(provider, 'cfg', { kind: 'ai', prompt: 'a' }).catch(() => undefined);
    await expect(engine.run(provider, 'cfg', { kind: 'ai', prompt: 'b' })).rejects.toThrow(/waited over 30 ms/);
    release();
    await running;
  });
});

describe('change-set retention (L3)', () => {
  const write = (cell: string): ChangeOp[] => [
    { kind: 'write_range', range: `Sheet1!${cell}`, cells: [[{ value: cell }]], reason: 'test' },
  ];

  it('keeps at most maxSets change-sets, discarding the oldest', async () => {
    const mgr = new ChangeSetManager(new WorkbookEmulator(['Sheet1']), 'wb', {
      ...DEFAULT_CHANGESET_RETENTION,
      maxSets: 3,
    });
    const ids: string[] = [];
    for (const cell of ['A1', 'A2', 'A3', 'A4', 'A5']) ids.push((await mgr.stage(write(cell))).id);
    expect(mgr.list().map((cs) => cs.id)).toEqual(ids.slice(2));
    await expect(mgr.apply(ids[0]!)).rejects.toThrow(/Unknown change-set.*only recent change-sets are kept/);
  });

  it('discards change-sets older than maxAgeMs', async () => {
    vi.useFakeTimers({ toFake: ['Date'] });
    vi.setSystemTime(new Date('2026-01-01T00:00:00Z'));
    const mgr = new ChangeSetManager(new WorkbookEmulator(['Sheet1']), 'wb', {
      ...DEFAULT_CHANGESET_RETENTION,
      maxAgeMs: 60_000,
    });
    const old = await mgr.stage(write('A1'));
    vi.setSystemTime(new Date('2026-01-01T00:02:00Z'));
    const fresh = await mgr.stage(write('A2'));
    expect(mgr.get(old.id)).toBeUndefined();
    expect(mgr.get(fresh.id)).toBeDefined();
  });

  it('bounds the cells held in snapshots across all retained sets', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    // Each applied 10-cell write keeps a 10-cell snapshot and a 10-cell re-read.
    const mgr = new ChangeSetManager(wb, 'wb', { ...DEFAULT_CHANGESET_RETENTION, maxRetainedCells: 45 });
    const ids: string[] = [];
    for (const col of ['A', 'B', 'C']) {
      const cs = await mgr.stage([
        {
          kind: 'write_range',
          range: `Sheet1!${col}1:${col}10`,
          cells: Array.from({ length: 10 }, () => [{ value: 1 }]),
          reason: 'fill',
        },
      ]);
      ids.push(cs.id);
      expect((await mgr.apply(cs.id)).ok).toBe(true);
    }
    // 3 × 20 = 60 cells > 45, so the oldest set was dropped once the third settled.
    expect(mgr.get(ids[0]!)).toBeUndefined();
    expect(mgr.get(ids[1]!)?.status).toBe('applied');
    expect(mgr.get(ids[2]!)?.status).toBe('applied');
  });

  it('never discards a set that is mid-apply', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const mgr = new ChangeSetManager(wb, 'wb', { ...DEFAULT_CHANGESET_RETENTION, maxSets: 1 });
    const first = await mgr.stage(write('A1'));
    const applying = mgr.apply(first.id); // status flips to 'applying' synchronously
    const second = await mgr.stage(write('A2'));
    expect(mgr.get(first.id)).toBeDefined(); // over budget, but busy
    await applying;
    // Once it settles, the budget is enforced again.
    expect(mgr.get(first.id)).toBeUndefined();
    expect(mgr.get(second.id)).toBeDefined();
  });
});
