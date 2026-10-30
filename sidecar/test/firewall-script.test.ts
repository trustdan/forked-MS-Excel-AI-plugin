import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { isIPv4, isIPv6 } from 'node:net';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';

const scriptPath = join(__dirname, '..', '..', 'scripts', 'lockdown-firewall.ps1');
const source = readFileSync(scriptPath, 'utf8');

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
});
