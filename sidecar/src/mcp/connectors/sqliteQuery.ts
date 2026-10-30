import { spawn } from 'node:child_process';

export const DEFAULT_QUERY_TIMEOUT_MS = 10_000;
export const MAX_RESULT_BYTES = 4 * 1024 * 1024;

export interface BoundedQueryResult {
  rows: Record<string, unknown>[];
  /** True when the row or byte cap stopped the query before it ran out of rows. */
  truncated: boolean;
}

// Runs in a child process. node:sqlite has no interrupt or progress handler, and a
// worker thread cannot be stopped mid-step, so killing a process is the only way to
// end a runaway query. The caps are enforced by stepping rows, never by editing SQL.
const RUNNER = `
const { DatabaseSync } = require('node:sqlite');
let input = '';
process.stdin.setEncoding('utf8');
process.stdin.on('data', (c) => (input += c));
process.stdin.on('end', () => {
  const { dbPath, sql, maxRows, maxBytes } = JSON.parse(input);
  let out;
  try {
    const db = new DatabaseSync(dbPath, { readOnly: true });
    const rows = [];
    let bytes = 0;
    let truncated = false;
    for (const raw of db.prepare(sql).iterate()) {
      if (rows.length === maxRows) { truncated = true; break; }
      const row = {};
      for (const [k, v] of Object.entries(raw)) {
        row[k] = v instanceof Uint8Array ? '[blob: ' + v.length + ' bytes]' : v;
      }
      bytes += JSON.stringify(row).length;
      if (bytes > maxBytes) { truncated = true; break; }
      rows.push(row);
    }
    db.close();
    out = { ok: true, rows, truncated };
  } catch (e) {
    out = { ok: false, error: e instanceof Error ? e.message : String(e) };
  }
  process.stdout.write(JSON.stringify(out));
});
`;

/**
 * Runs one read-only statement against `dbPath` in a child process, returning at
 * most `maxRows` rows and about `MAX_RESULT_BYTES` of data. The child is killed when
 * `timeoutMs` passes or `abortSignal` fires, so a slow query never blocks the sidecar.
 */
export function runBoundedQuery(args: {
  dbPath: string;
  sql: string;
  maxRows: number;
  timeoutMs: number;
  abortSignal?: AbortSignal;
}): Promise<BoundedQueryResult> {
  const { dbPath, sql, maxRows, timeoutMs, abortSignal } = args;
  return new Promise((resolve, reject) => {
    if (abortSignal?.aborted) {
      reject(new Error('Query stopped: the request was cancelled.'));
      return;
    }
    const child = spawn(process.execPath, ['--no-warnings', '-e', RUNNER], {
      stdio: ['pipe', 'pipe', 'ignore'],
      windowsHide: true,
    });
    let stdout = '';
    let settled = false;
    const finish = (err: Error | null, result?: BoundedQueryResult): void => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      abortSignal?.removeEventListener('abort', onAbort);
      if (child.exitCode === null) child.kill();
      if (err) reject(err);
      else resolve(result!);
    };
    const onAbort = (): void => finish(new Error('Query stopped: the request was cancelled.'));
    const timer = setTimeout(
      () => finish(new Error(`Query stopped: it exceeded the ${timeoutMs / 1000}s execution budget.`)),
      timeoutMs,
    );
    abortSignal?.addEventListener('abort', onAbort, { once: true });

    child.stdout.setEncoding('utf8');
    child.stdout.on('data', (chunk: string) => {
      stdout += chunk;
      // Row bytes are capped in the child; this only guards against a misbehaving one.
      if (stdout.length > MAX_RESULT_BYTES * 2) finish(new Error('Query result exceeded the output limit.'));
    });
    child.on('error', (e) => finish(e));
    child.stdin.on('error', () => {}); // a child that died early is reported by 'close'
    child.on('close', () => {
      try {
        const out = JSON.parse(stdout) as
          | { ok: true; rows: Record<string, unknown>[]; truncated: boolean }
          | { ok: false; error: string };
        if (out.ok) finish(null, { rows: out.rows, truncated: out.truncated });
        else finish(new Error(out.error));
      } catch {
        finish(new Error('Query process exited without a result.'));
      }
    });
    child.stdin.end(JSON.stringify({ dbPath, sql, maxRows, maxBytes: MAX_RESULT_BYTES }));
  });
}
