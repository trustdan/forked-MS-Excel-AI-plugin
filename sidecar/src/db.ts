import { DatabaseSync } from 'node:sqlite';
import { join } from 'node:path';
import { dataDir } from './config.js';

export type Db = DatabaseSync;

const MIGRATIONS = `
CREATE TABLE IF NOT EXISTS tokens (
  token TEXT PRIMARY KEY,
  created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS fn_cache (
  key TEXT PRIMARY KEY,
  value TEXT NOT NULL,
  created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS ledger (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  at TEXT NOT NULL,
  provider TEXT NOT NULL,
  model TEXT NOT NULL,
  feature TEXT NOT NULL,
  input_tokens INTEGER NOT NULL,
  output_tokens INTEGER NOT NULL
);
CREATE INDEX IF NOT EXISTS fn_cache_created_at ON fn_cache(created_at);
CREATE INDEX IF NOT EXISTS ledger_at ON ledger(at);
CREATE TABLE IF NOT EXISTS memory (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  scope TEXT NOT NULL,            -- 'user' or workbook id
  content TEXT NOT NULL,
  created_at TEXT NOT NULL
);
CREATE VIRTUAL TABLE IF NOT EXISTS memory_fts USING fts5(
  content,
  content='memory',
  content_rowid='id'
);
CREATE TRIGGER IF NOT EXISTS memory_ai AFTER INSERT ON memory BEGIN
  INSERT INTO memory_fts(rowid, content) VALUES (new.id, new.content);
END;
CREATE TRIGGER IF NOT EXISTS memory_ad AFTER DELETE ON memory BEGIN
  INSERT INTO memory_fts(memory_fts, rowid, content) VALUES ('delete', old.id, old.content);
END;
`;

/**
 * Lifetime budgets for the tables that grow with use. Each table is pruned
 * after every write to it and once at open, so neither recalc storms nor a
 * long-lived install can grow the database without bound.
 */
export interface RetentionLimits {
  fnCacheMaxRows: number;
  fnCacheMaxAgeMs: number;
  ledgerMaxRows: number;
  ledgerMaxAgeMs: number;
}

const DAY_MS = 24 * 60 * 60 * 1000;

export const DEFAULT_RETENTION: RetentionLimits = {
  // The age limit also bounds staleness: an entry is recomputed at least this
  // often, so a model swapped behind an unchanged name and endpoint is seen.
  fnCacheMaxRows: 20_000,
  fnCacheMaxAgeMs: 30 * DAY_MS,
  ledgerMaxRows: 200_000,
  ledgerMaxAgeMs: 400 * DAY_MS,
};

export function pruneFnCache(db: Db, limits: RetentionLimits = DEFAULT_RETENTION, now = Date.now()): void {
  db.prepare('DELETE FROM fn_cache WHERE created_at < ?').run(new Date(now - limits.fnCacheMaxAgeMs).toISOString());
  db.prepare(
    'DELETE FROM fn_cache WHERE key IN (SELECT key FROM fn_cache ORDER BY created_at DESC LIMIT -1 OFFSET ?)',
  ).run(limits.fnCacheMaxRows);
}

export function pruneLedger(db: Db, limits: RetentionLimits = DEFAULT_RETENTION, now = Date.now()): void {
  db.prepare('DELETE FROM ledger WHERE at < ?').run(new Date(now - limits.ledgerMaxAgeMs).toISOString());
  db.prepare('DELETE FROM ledger WHERE id IN (SELECT id FROM ledger ORDER BY id DESC LIMIT -1 OFFSET ?)').run(
    limits.ledgerMaxRows,
  );
}

export function openDb(path?: string, limits: RetentionLimits = DEFAULT_RETENTION): Db {
  const db = new DatabaseSync(path ?? join(dataDir(), 'excelai.db'));
  db.exec('PRAGMA journal_mode = WAL;');
  db.exec(MIGRATIONS);
  pruneFnCache(db, limits);
  pruneLedger(db, limits);
  return db;
}
