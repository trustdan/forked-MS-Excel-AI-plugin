import { request as httpRequest, type IncomingMessage } from 'node:http';
import { request as httpsRequest } from 'node:https';

/**
 * Loopback Network Guard
 *
 * Enforces strict air-gapped local execution. Under no circumstances may
 * workbook data, prompts, or model requests leave the local host.
 */

export class NonLocalConnectionBlockedError extends Error {
  constructor(attemptedUrl: string) {
    super(
      `[Air-Gap Security] Outbound connection blocked: "${attemptedUrl}" is not a permitted loopback address (127.0.0.1, localhost, [::1]). Zero data may leave localhost.`,
    );
    this.name = 'NonLocalConnectionBlockedError';
  }
}

const LOOPBACK_HOSTNAMES = new Set(['localhost', '127.0.0.1', '[::1]', '::1']);
const INITIAL_NATIVE_FETCH = globalThis.fetch;

/**
 * Validates that a target URL is strictly hosted on loopback (127.0.0.0/8, localhost, [::1]).
 * Throws NonLocalConnectionBlockedError if any external or non-loopback network destination is targeted.
 */
export function assertLoopbackUrl(urlInput: string | URL): URL {
  let parsed: URL;
  try {
    parsed = typeof urlInput === 'string' ? new URL(urlInput) : urlInput;
  } catch {
    throw new NonLocalConnectionBlockedError(String(urlInput));
  }

  const hostname = parsed.hostname.toLowerCase();

  // Allow standard loopback hostnames
  if (LOOPBACK_HOSTNAMES.has(hostname)) {
    return parsed;
  }

  // Allow any IPv4 127.0.0.0/8 address (e.g. 127.0.0.2)
  if (/^127(?:\.(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)){3}$/.test(hostname)) {
    return parsed;
  }

  throw new NonLocalConnectionBlockedError(parsed.href);
}

/**
 * Loopback-only HTTP client using `node:http` / `node:https` without undici's
 * 300-second `headersTimeout` (which otherwise aborts long local LLM generations
 * with `TypeError: fetch failed` / `[LM STUDIO SERVER] Client disconnected`).
 *
 * Automatically delegates to `globalThis.fetch` when `fetch` is mocked in tests.
 */
export async function loopbackFetch(url: string, init: RequestInit = {}): Promise<Response> {
  const parsed = assertLoopbackUrl(url);

  if (globalThis.fetch !== INITIAL_NATIVE_FETCH) {
    return globalThis.fetch(url, init);
  }

  const signal = init.signal ?? undefined;
  if (signal?.aborted) {
    const abortErr = new Error('The operation was aborted');
    abortErr.name = 'AbortError';
    throw abortErr;
  }

  const isHttps = parsed.protocol === 'https:';
  const reqFn = isHttps ? httpsRequest : httpRequest;
  const bodyStr = typeof init.body === 'string' ? init.body : undefined;
  const bodyBuf = bodyStr !== undefined ? Buffer.from(bodyStr, 'utf8') : undefined;

  const headers: Record<string, string> = {};
  if (init.headers) {
    if (init.headers instanceof Headers) {
      init.headers.forEach((v, k) => {
        headers[k] = v;
      });
    } else if (Array.isArray(init.headers)) {
      for (const [k, v] of init.headers) {
        headers[k] = v;
      }
    } else {
      Object.assign(headers, init.headers);
    }
  }
  if (bodyBuf !== undefined && !('content-length' in headers) && !('Content-Length' in headers)) {
    headers['content-length'] = String(bodyBuf.byteLength);
  }

  return new Promise<Response>((resolve, reject) => {
    let settled = false;
    const cleanup = (): void => {
      if (signal) {
        signal.removeEventListener('abort', onAbort);
      }
    };

    const fail = (err: Error): void => {
      if (settled) return;
      settled = true;
      cleanup();
      reject(err);
    };

    const req = reqFn(
      parsed,
      {
        method: init.method ?? 'GET',
        headers,
        timeout: 0,
      },
      (res: IncomingMessage) => {
        const chunks: Buffer[] = [];
        res.on('data', (chunk: Buffer) => {
          chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
        });
        res.on('end', () => {
          if (settled) return;
          settled = true;
          cleanup();
          const resHeaders = new Headers();
          for (const [k, v] of Object.entries(res.headers)) {
            if (typeof v === 'string') resHeaders.set(k, v);
            else if (Array.isArray(v)) resHeaders.set(k, v.join(', '));
          }
          resolve(
            new Response(Buffer.concat(chunks), {
              status: res.statusCode ?? 500,
              statusText: res.statusMessage ?? '',
              headers: resHeaders,
            }),
          );
        });
        res.on('error', (err) => {
          if (signal?.aborted) {
            const abortErr = new Error('The operation was aborted');
            abortErr.name = 'AbortError';
            fail(abortErr);
            return;
          }
          fail(formatLoopbackSocketError(parsed, err));
        });
      },
    );

    const onAbort = (): void => {
      req.destroy();
      const abortErr = new Error('The operation was aborted');
      abortErr.name = 'AbortError';
      fail(abortErr);
    };

    if (signal) {
      signal.addEventListener('abort', onAbort, { once: true });
    }

    req.on('socket', (socket) => {
      socket.setNoDelay(true);
      socket.setKeepAlive(true, 15_000);
      socket.setTimeout(0);
    });

    req.on('error', (err) => {
      if (signal?.aborted) {
        const abortErr = new Error('The operation was aborted');
        abortErr.name = 'AbortError';
        fail(abortErr);
        return;
      }
      fail(formatLoopbackSocketError(parsed, err));
    });

    if (bodyBuf !== undefined) {
      req.write(bodyBuf);
    }
    req.end();
  });
}

function formatLoopbackSocketError(url: URL, err: Error & { code?: string }): Error {
  const code = err.code ?? '';
  if (code === 'ECONNREFUSED') {
    return new Error(
      `Cannot connect to local model server at ${url.origin} (ECONNREFUSED). Make sure LM Studio or Ollama is running and its local server is started.`,
    );
  }
  if (code === 'ECONNRESET' || err.message.toLowerCase().includes('socket hang up')) {
    return new Error(
      `Local model server at ${url.origin} closed the connection unexpectedly (${code || 'socket hang up'}). Check LM Studio to ensure the model did not run out of memory or context length.`,
    );
  }
  return new Error(`Local model request to ${url.origin} failed: ${err.message}`);
}
