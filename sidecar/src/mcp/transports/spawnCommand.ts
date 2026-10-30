import { existsSync, statSync } from 'node:fs';
import { win32 } from 'node:path';

const { delimiter, extname, isAbsolute, join } = win32;

export interface SpawnPlan {
  file: string;
  args: string[];
  /** Pass the command line to cmd.exe exactly as built (Windows only). */
  windowsVerbatimArguments?: boolean;
}

/**
 * Characters cmd.exe still interprets inside a double-quoted argument. `"`
 * ends the quoting and `%` expands variables even when quoted (the BatBadBut
 * class of injection), and a line break ends the command. There is no reliable
 * escape for these through a batch shim, so they are refused.
 */
const CMD_UNSAFE = /["%\r\n\0]/;

/**
 * Decide how to start an MCP server command without `shell: true`.
 *
 * POSIX: spawn the command directly. Windows: resolve it against PATH and
 * PATHEXT. A `.exe`/`.com` is spawned directly, and Node quotes its arguments.
 * A `.cmd`/`.bat` (such as `npx`) can only run through cmd.exe, so every
 * argument is quoted and arguments cmd.exe could reinterpret are rejected.
 */
export function planSpawn(
  command: string,
  args: readonly string[],
  env: NodeJS.ProcessEnv,
  platform: NodeJS.Platform = process.platform,
): SpawnPlan {
  if (platform !== 'win32') return { file: command, args: [...args] };

  const resolved = resolveWindowsCommand(command, env) ?? command;
  const ext = extname(resolved).toLowerCase();
  if (ext !== '.cmd' && ext !== '.bat') return { file: resolved, args: [...args] };

  for (const part of [resolved, ...args]) {
    if (CMD_UNSAFE.test(part)) {
      throw new Error(
        `MCP command "${command}" is a batch file, and an argument contains a character cmd.exe would ` +
          `reinterpret (", %, or a line break). Point the server at the .exe (for example node.exe and the ` +
          `script path) instead of a .cmd/.bat wrapper.`,
      );
    }
  }
  const quoted = [resolved, ...args].map(quoteForCmd).join(' ');
  return {
    file: envValue(env, 'ComSpec') ?? 'cmd.exe',
    // /d skips AutoRun, /v:off keeps ! literal, /s strips only the outer quotes.
    args: ['/d', '/v:off', '/s', '/c', `"${quoted}"`],
    windowsVerbatimArguments: true,
  };
}

/** Quote for both cmd.exe and the MSVCRT argv parser of the program it runs. */
function quoteForCmd(arg: string): string {
  // Backslashes before a quote are escapes for MSVCRT; double the trailing run.
  return `"${arg.replace(/(\\+)$/, '$1$1')}"`;
}

/** Find what CreateProcess-via-cmd would run, without searching the cwd. */
export function resolveWindowsCommand(command: string, env: NodeJS.ProcessEnv): string | null {
  const exts = (envValue(env, 'PATHEXT') ?? '.COM;.EXE;.BAT;.CMD')
    .split(';')
    .filter(Boolean)
    .map((e) => e.toLowerCase());
  const candidates = (base: string): string[] =>
    extname(base) && exts.includes(extname(base).toLowerCase()) ? [base] : exts.map((e) => base + e);

  const hasDir = isAbsolute(command) || /[\\/]/.test(command);
  const bases = hasDir
    ? [win32.resolve(command)]
    : (envValue(env, 'PATH') ?? '').split(delimiter).filter(Boolean).map((dir) => join(dir, command));

  for (const base of bases) {
    for (const file of candidates(base)) {
      try {
        if (existsSync(file) && statSync(file).isFile()) return file;
      } catch {
        // unreadable entry; keep looking
      }
    }
  }
  return null;
}

/** Windows environment names are case-insensitive, but a copied env object is not. */
function envValue(env: NodeJS.ProcessEnv, name: string): string | undefined {
  const key = Object.keys(env).find((k) => k.toLowerCase() === name.toLowerCase());
  return key ? env[key] : undefined;
}
