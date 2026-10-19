import { resolve, normalize, relative, isAbsolute } from 'node:path';
import { existsSync, realpathSync } from 'node:fs';

export class McpSecurityError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'McpSecurityError';
  }
}

/**
 * PathGuard strictly restricts all local filesystem, SQLite, and CSV operations
 * to an explicit list of user-approved directory and file paths.
 * Enforces zero-traversal and symlink verification.
 */
export class PathGuard {
  /**
   * Resolve and canonicalize path, expanding symlinks if the path exists.
   */
  static canonicalize(targetPath: string): string {
    if (!targetPath || typeof targetPath !== 'string') {
      throw new McpSecurityError('Path must be a non-empty string');
    }
    if (targetPath.includes('\0')) {
      throw new McpSecurityError('Null bytes in path are prohibited');
    }
    const resolved = resolve(targetPath);
    if (existsSync(resolved)) {
      try {
        return realpathSync(resolved);
      } catch {
        return normalize(resolved);
      }
    }
    return normalize(resolved);
  }

  /**
   * Determine whether a target path lies within or matches one of the approved paths.
   */
  static isApproved(targetPath: string, approvedPaths: string[]): boolean {
    if (!approvedPaths || approvedPaths.length === 0) {
      return false;
    }
    try {
      const canonicalTarget = this.canonicalize(targetPath);
      // Case-insensitive comparison on Windows, case-sensitive on POSIX
      const isWindows = process.platform === 'win32';

      for (const approved of approvedPaths) {
        if (!approved) continue;
        const canonicalApproved = this.canonicalize(approved);

        if (isWindows) {
          const lowerTarget = canonicalTarget.toLowerCase();
          const lowerApproved = canonicalApproved.toLowerCase();
          if (lowerTarget === lowerApproved) return true;
          const rel = relative(canonicalApproved, canonicalTarget);
          if (!rel.startsWith('..') && !isAbsolute(rel) && !lowerTarget.startsWith('..')) {
            // Check that it's a true subdirectory/child
            const relLower = rel.toLowerCase();
            if (!relLower.startsWith('..')) {
              return true;
            }
          }
        } else {
          if (canonicalTarget === canonicalApproved) return true;
          const rel = relative(canonicalApproved, canonicalTarget);
          if (!rel.startsWith('..') && !isAbsolute(rel)) {
            return true;
          }
        }
      }
      return false;
    } catch {
      return false;
    }
  }

  /**
   * Assert that targetPath is within approvedPaths, throwing McpSecurityError if not.
   */
  static assertApproved(targetPath: string, approvedPaths: string[], context?: string): string {
    const canonical = this.canonicalize(targetPath);
    if (!this.isApproved(canonical, approvedPaths)) {
      const msg = context ? ` (${context})` : '';
      throw new McpSecurityError(
        `SecurityError: Access denied${msg}. Path "${targetPath}" is outside approved paths: [${approvedPaths.join(', ')}]`,
      );
    }
    return canonical;
  }
}
