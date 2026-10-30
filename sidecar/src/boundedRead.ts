import { closeSync, fstatSync, openSync, readSync } from 'node:fs';

/**
 * Reads at most `cap` bytes through one open handle, so memory follows the cap
 * rather than the file's size, and the size and bytes come from the same file
 * even if the path is swapped between the check and the read.
 */
export function readFileCapped(path: string, cap: number): { bytes: Buffer; size: number; truncated: boolean } {
  const fd = openSync(path, 'r');
  try {
    const stat = fstatSync(fd);
    if (!stat.isFile()) throw new Error(`Path is not a regular file: "${path}"`);
    // One byte past the cap tells a file that fits from one that was cut short, even if it grew after fstat.
    const buf = Buffer.alloc(Math.min(cap + 1, stat.size + 1));
    let filled = 0;
    while (filled < buf.length) {
      const n = readSync(fd, buf, filled, buf.length - filled, filled);
      if (n === 0) break;
      filled += n;
    }
    const truncated = filled > cap;
    return { bytes: buf.subarray(0, Math.min(filled, cap)), size: stat.size, truncated };
  } finally {
    closeSync(fd);
  }
}
