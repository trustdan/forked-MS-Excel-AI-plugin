import { describe, it, expect } from 'vitest';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { mkdirSync, writeFileSync, rmSync } from 'node:fs';
import { PathGuard, McpSecurityError } from '../src/mcp/pathGuard.js';

describe('PathGuard', () => {
  const baseDir = join(tmpdir(), `excelai-guard-test-${Date.now()}`);
  const approvedDir = join(baseDir, 'approved');
  const secretDir = join(baseDir, 'secret');

  mkdirSync(approvedDir, { recursive: true });
  mkdirSync(secretDir, { recursive: true });
  writeFileSync(join(approvedDir, 'data.csv'), 'a,b,c');
  writeFileSync(join(secretDir, 'passwords.txt'), 'secret');

  const approvedList = [approvedDir];

  it('approves exact directory and files inside approved path', () => {
    expect(PathGuard.isApproved(approvedDir, approvedList)).toBe(true);
    expect(PathGuard.isApproved(join(approvedDir, 'data.csv'), approvedList)).toBe(true);
    expect(() => PathGuard.assertApproved(join(approvedDir, 'data.csv'), approvedList)).not.toThrow();
  });

  it('approves nested subdirectories within approved path', () => {
    const sub = join(approvedDir, 'sub', 'folder');
    mkdirSync(sub, { recursive: true });
    expect(PathGuard.isApproved(sub, approvedList)).toBe(true);
    expect(PathGuard.isApproved(join(sub, 'nested.txt'), approvedList)).toBe(true);
  });

  it('rejects paths outside approved paths with McpSecurityError', () => {
    expect(PathGuard.isApproved(secretDir, approvedList)).toBe(false);
    expect(PathGuard.isApproved(join(secretDir, 'passwords.txt'), approvedList)).toBe(false);

    expect(() => PathGuard.assertApproved(join(secretDir, 'passwords.txt'), approvedList)).toThrow(
      McpSecurityError,
    );
  });

  it('rejects directory traversal attacks (../)', () => {
    const traversal = join(approvedDir, '..', 'secret', 'passwords.txt');
    expect(PathGuard.isApproved(traversal, approvedList)).toBe(false);
    expect(() => PathGuard.assertApproved(traversal, approvedList)).toThrow(McpSecurityError);
  });

  it('rejects null bytes in path', () => {
    expect(() => PathGuard.assertApproved(approvedDir + '\0.csv', approvedList)).toThrow(
      McpSecurityError,
    );
  });

  it('rejects when approved list is empty', () => {
    expect(PathGuard.isApproved(join(approvedDir, 'data.csv'), [])).toBe(false);
    expect(() => PathGuard.assertApproved(join(approvedDir, 'data.csv'), [])).toThrow(
      McpSecurityError,
    );
  });
});
