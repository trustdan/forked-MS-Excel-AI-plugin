import { createServer } from 'node:http';
import type { AddressInfo } from 'node:net';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { guardedFetch, NonLocalConnectionBlockedError } from '../src/networkGuard.js';

// Simulates a hostile hosts file / resolver: the URL passes the hostname check,
// but what it resolves to decides where the socket may connect.
const resolved = vi.hoisted(() => ({ addresses: [] as { address: string; family: number }[] }));
vi.mock('node:dns/promises', () => ({
  lookup: async () => resolved.addresses,
}));

const servers: ReturnType<typeof createServer>[] = [];
afterEach(async () => {
  for (const s of servers.splice(0)) await new Promise<void>((r) => s.close(() => r()));
});

describe('guardedFetch pins the resolved address', () => {
  it('blocks `localhost` when it resolves only to a non-loopback address', async () => {
    resolved.addresses = [{ address: '10.0.0.5', family: 4 }];
    await expect(guardedFetch('http://localhost:1234/v1/models')).rejects.toBeInstanceOf(
      NonLocalConnectionBlockedError,
    );
  });

  it('connects only to the loopback entries when resolution is mixed', async () => {
    const server = createServer((_req, res) => res.end('pinned'));
    servers.push(server);
    await new Promise<void>((r) => server.listen(0, '127.0.0.1', r));
    const { port } = server.address() as AddressInfo;

    resolved.addresses = [
      { address: '192.0.2.1', family: 4 },
      { address: '127.0.0.1', family: 4 },
    ];
    const res = await guardedFetch(`http://localhost:${port}/`);
    expect(await res.text()).toBe('pinned');
  });
});
