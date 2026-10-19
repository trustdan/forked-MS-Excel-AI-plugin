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
  } catch (err) {
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
