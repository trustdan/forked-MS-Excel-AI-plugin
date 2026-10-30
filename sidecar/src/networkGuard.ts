import { lookup } from 'node:dns/promises';
import type { LookupAddress } from 'node:dns';
import { request as httpRequest, type IncomingMessage } from 'node:http';
import { request as httpsRequest } from 'node:https';
import { isIP, type LookupFunction } from 'node:net';
import { ProviderError } from './providers/types.js';

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

export class RedirectBlockedError extends Error {
  constructor(from: string, location: string) {
    super(
      `[Air-Gap Security] Redirect from "${from}" to "${location}" blocked: outbound requests never follow redirects.`,
    );
    this.name = 'RedirectBlockedError';
  }
}

/** Response bodies above this are rejected while being read, never fully retained. */
export const DEFAULT_MAX_RESPONSE_BYTES = 8 * 1024 * 1024; // 8 MiB

/** Socket idle time allowed once headers arrive. Before headers there is no limit: local generations can be slow. */
export const DEFAULT_INACTIVITY_TIMEOUT_MS = 180_000; // 3 minutes

export interface GuardedRequestInit {
  method?: string;
  headers?: Record<string, string>;
  body?: string;
  signal?: AbortSignal | null;
  maxResponseBytes?: number;
  inactivityTimeoutMs?: number;
}

/** Outbound HTTP. Production uses `guardedFetch`; tests inject a stub. */
export type Transport = (url: string, init?: GuardedRequestInit) => Promise<Response>;

/** True for 127.0.0.0/8 and ::1. IPv4-mapped IPv6 (::ffff:127.0.0.1) is deliberately not accepted. */
export function isLoopbackAddress(address: string): boolean {
  if (isIP(address) === 4) return address.startsWith('127.');
  return isIP(address) === 6 && /^(?:0{0,4}:){2,7}0{0,3}1$|^::1$/.test(address);
}

/** Resolves the URL's host to the loopback addresses it may connect to; throws if there are none. */
async function resolveLoopback(url: URL): Promise<LookupAddress[]> {
  const host = url.hostname.replace(/^\[|\]$/g, '');
  const family = isIP(host);
  const addresses = family !== 0 ? [{ address: host, family }] : await lookup(host, { all: true });
  const allowed = addresses.filter((a) => isLoopbackAddress(a.address));
  if (allowed.length === 0) throw new NonLocalConnectionBlockedError(url.href);
  return allowed;
}

/**
 * The single outbound HTTP path. The host is resolved and checked before
 * connecting, the socket may only dial those loopback addresses, and its actual
 * remote address is checked again on connect. Redirects are refused. The
 * response resolves at header time with a streaming body that enforces a byte
 * cap and an inactivity timeout on the socket; abort or cancel destroys it.
 *
 * Uses `node:http` rather than undici to avoid its 300-second `headersTimeout`,
 * which aborts long local LLM generations.
 */
export const guardedFetch: Transport = async (url, init = {}) => {
  const parsed = assertLoopbackUrl(url);
  const signal = init.signal ?? undefined;
  if (signal?.aborted) throw abortError();
  const addresses = await resolveLoopback(parsed);
  if (signal?.aborted) throw abortError();

  const maxBytes = init.maxResponseBytes ?? DEFAULT_MAX_RESPONSE_BYTES;
  const idleMs = init.inactivityTimeoutMs ?? DEFAULT_INACTIVITY_TIMEOUT_MS;
  const body = init.body !== undefined ? Buffer.from(init.body, 'utf8') : undefined;
  const headers = { ...init.headers };
  if (body && !Object.keys(headers).some((k) => k.toLowerCase() === 'content-length')) {
    headers['content-length'] = String(body.byteLength);
  }

  return new Promise<Response>((resolve, reject) => {
    // Recorded before destroying the request so every listener reports this cause, not a generic reset.
    let failure: Error | undefined;
    let onBodyError: ((err: Error) => void) | undefined;
    const kill = (err: Error): void => {
      failure ??= err;
      req.destroy(err);
    };

    const req = (parsed.protocol === 'https:' ? httpsRequest : httpRequest)(parsed, {
      method: init.method ?? 'GET',
      headers,
      agent: false,
      lookup: ((_host, opts, cb) => {
        if (opts.all) cb(null, addresses);
        else cb(null, addresses[0]!.address, addresses[0]!.family);
      }) as LookupFunction,
    });

    const onAbort = (): void => kill(abortError());
    signal?.addEventListener('abort', onAbort, { once: true });
    req.on('close', () => signal?.removeEventListener('abort', onAbort));

    req.on('socket', (socket) => {
      socket.setNoDelay(true);
      socket.once('connect', () => {
        if (!isLoopbackAddress(socket.remoteAddress ?? '')) {
          kill(new NonLocalConnectionBlockedError(parsed.href));
        }
      });
    });

    req.on('error', (err) => {
      const cause = failure ?? formatLoopbackSocketError(parsed, err);
      if (onBodyError) onBodyError(cause);
      else reject(cause);
    });

    req.on('response', (res: IncomingMessage) => {
      const status = res.statusCode ?? 500;
      const location = res.headers.location;
      if (status >= 300 && status < 400 && location) {
        kill(new RedirectBlockedError(parsed.href, location));
        return;
      }
      const declared = Number(res.headers['content-length']);
      if (Number.isFinite(declared) && declared > maxBytes) {
        kill(tooLargeError(maxBytes));
        return;
      }
      if (idleMs > 0) {
        res.socket.setTimeout(idleMs, () =>
          kill(
            new ProviderError(
              `Local model server at ${parsed.origin} sent nothing for ${Math.round(idleMs / 1000)}s. Check if it is stuck or out of context memory.`,
              504,
              false,
            ),
          ),
        );
      }

      let finished = false;
      let received = 0;
      const stream = new ReadableStream<Uint8Array>({
        start(controller) {
          onBodyError = (err) => {
            if (finished) return;
            finished = true;
            controller.error(err);
          };
          res.on('data', (chunk: Buffer) => {
            if (finished) return;
            received += chunk.byteLength;
            if (received > maxBytes) {
              kill(tooLargeError(maxBytes));
              return;
            }
            controller.enqueue(new Uint8Array(chunk.buffer, chunk.byteOffset, chunk.byteLength));
            if ((controller.desiredSize ?? 0) <= 0) res.pause();
          });
          res.on('end', () => {
            if (finished) return;
            finished = true;
            res.socket?.setTimeout(0);
            controller.close();
          });
          res.on('error', (err) => onBodyError?.(failure ?? err));
          res.on('close', () => {
            if (!res.complete) {
              onBodyError?.(
                failure ?? new Error(`Connection to ${parsed.origin} closed before the response finished`),
              );
            }
          });
        },
        pull() {
          res.resume();
        },
        cancel() {
          finished = true;
          req.destroy();
        },
      });

      const resHeaders = new Headers();
      for (const [k, v] of Object.entries(res.headers)) {
        if (typeof v === 'string') resHeaders.set(k, v);
        else if (Array.isArray(v)) resHeaders.set(k, v.join(', '));
      }
      const nullBody = status === 204 || status === 205 || status === 304;
      if (nullBody) res.resume();
      resolve(
        new Response(nullBody ? null : stream, {
          status,
          statusText: res.statusMessage ?? '',
          headers: resHeaders,
        }),
      );
    });

    if (body !== undefined) req.write(body);
    req.end();
  });
};

function abortError(): Error {
  const err = new Error('The operation was aborted');
  err.name = 'AbortError';
  return err;
}

function tooLargeError(maxBytes: number): ProviderError {
  return new ProviderError(`Response exceeded maximum allowed size (${maxBytes} bytes)`, 413, false);
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
