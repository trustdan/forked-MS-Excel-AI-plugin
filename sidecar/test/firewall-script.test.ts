import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { isIPv4, isIPv6 } from 'node:net';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';

const repoRoot = join(__dirname, '..', '..');
const scriptPath = join(repoRoot, 'scripts', 'lockdown-firewall.ps1');
const nodeOfflinePath = join(repoRoot, 'bin', 'node-offline.exe');
const source = readFileSync(scriptPath, 'utf8');

function hasInstalledFirewallRule(): boolean {
  if (process.platform !== 'win32' || !existsSync(nodeOfflinePath)) return false;
  try {
    const out = execFileSync(
      'powershell.exe',
      [
        '-NoProfile',
        '-NonInteractive',
        '-Command',
        "(Get-NetFirewallRule -DisplayName 'ExcelAI-Node-Offline-Block' -ErrorAction SilentlyContinue | Measure-Object).Count",
      ],
      { encoding: 'utf8' },
    );
    return Number(out.trim()) > 0;
  } catch {
    return false;
  }
}

function blockedRanges(): Array<[string, string]> {
  const block = /\$BlockedRanges = @\(([\s\S]*?)\)/.exec(source)?.[1] ?? '';
  return [...block.matchAll(/"([^"]+)"/g)].map((m) => {
    const [from, to] = m[1]!.split('-');
    return [from!, to ?? from!];
  });
}

function toBigInt(ip: string): bigint {
  if (isIPv4(ip)) return ip.split('.').reduce((acc, octet) => (acc << 8n) | BigInt(octet), 0n);
  const [head = '', tail = ''] = ip.split('::');
  const h = head ? head.split(':') : [];
  const t = tail ? tail.split(':') : [];
  const groups = ip.includes('::') ? [...h, ...Array(8 - h.length - t.length).fill('0'), ...t] : h;
  return groups.reduce((acc, g) => (acc << 16n) | BigInt(parseInt(g, 16)), 0n);
}

/** Gaps in [min, max] that no range covers, after merging. */
function uncovered(ranges: Array<[bigint, bigint]>, min: bigint, max: bigint): Array<[bigint, bigint]> {
  const gaps: Array<[bigint, bigint]> = [];
  let next = min;
  for (const [from, to] of [...ranges].sort((a, b) => (a[0] < b[0] ? -1 : 1))) {
    if (from > next) gaps.push([next, from - 1n]);
    if (to + 1n > next) next = to + 1n;
  }
  if (next <= max) gaps.push([next, max]);
  return gaps;
}

describe('lockdown-firewall.ps1 (L4)', () => {
  const ranges = blockedRanges();

  it('blocks every IPv4 address except 127.0.0.0/8', () => {
    const v4 = ranges.filter(([a]) => isIPv4(a)).map(([a, b]) => [toBigInt(a), toBigInt(b)] as [bigint, bigint]);
    expect(uncovered(v4, 0n, 0xffffffffn)).toEqual([[toBigInt('127.0.0.0'), toBigInt('127.255.255.255')]]);
  });

  it('blocks every IPv6 address except ::1 (and the unspecified ::)', () => {
    const v6 = ranges.filter(([a]) => isIPv6(a)).map(([a, b]) => [toBigInt(a), toBigInt(b)] as [bigint, bigint]);
    expect(v6.length).toBeGreaterThan(0);
    expect(uncovered(v6, 0n, (1n << 128n) - 1n)).toEqual([[0n, 1n]]);
  });

  it('is ASCII, so Windows PowerShell 5.1 reads a BOM-less file correctly', () => {
    expect([...source].filter((c) => c.charCodeAt(0) > 127)).toEqual([]);
  });

  it.skipIf(process.platform !== 'win32')('parses under Windows PowerShell 5.1', () => {
    const out = execFileSync(
      'powershell.exe',
      [
        '-NoProfile',
        '-NonInteractive',
        '-Command',
        `$e=$null; [void][System.Management.Automation.Language.Parser]::ParseFile('${scriptPath}',[ref]$null,[ref]$e); $e.Count`,
      ],
      { encoding: 'utf8' },
    );
    expect(out.trim()).toBe('0');
  });

  it('says that child MCP binaries need their own rules', () => {
    expect(source).toMatch(/MCP server[\s\S]*needs its own outbound block rule/);
  });

  it.skipIf(!hasInstalledFirewallRule())(
    'enforces loopback-only IPv4 and IPv6 on bin/node-offline.exe when the rule is installed',
    () => {
      const probeScript = `
        import net from 'node:net';
        async function listen(host) {
          const srv = net.createServer((s) => { s.on('error', () => {}); s.end('ok'); });
          await new Promise((res, rej) => { srv.once('error', rej); srv.listen(0, host, res); });
          return srv;
        }
        function probe(host, port) {
          return new Promise((res) => {
            const sock = net.connect({ host, port });
            const t = setTimeout(() => { sock.destroy(); res('TIMEOUT'); }, 1500);
            sock.once('connect', () => { clearTimeout(t); sock.end(); res('CONNECTED'); });
            sock.once('error', (err) => { clearTimeout(t); res(err.code || err.message); });
          });
        }
        const s4 = await listen('127.0.0.1');
        const s6 = await listen('::1');
        const out = {
          loopbackV4: await probe('127.0.0.1', s4.address().port),
          loopbackV6: await probe('::1', s6.address().port),
          externalV4: await probe('1.1.1.1', 443),
          externalV6: await probe('2606:4700:4700::1111', 443),
        };
        s4.close();
        s6.close();
        console.log(JSON.stringify(out));
      `;
      const out = execFileSync(nodeOfflinePath, ['--input-type=module', '-e', probeScript], {
        encoding: 'utf8',
      });
      expect(JSON.parse(out.trim())).toEqual({
        loopbackV4: 'CONNECTED',
        loopbackV6: 'CONNECTED',
        externalV4: 'EACCES',
        externalV6: 'EACCES',
      });
    },
  );
});
