import { existsSync } from 'node:fs';
import { spawn } from 'node:child_process';
import { join, resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const projectRoot = resolve(__dirname, '..');
const offlineNode = join(projectRoot, 'bin', 'node-offline.exe');
const sidecarEntry = join(projectRoot, 'sidecar', 'dist', 'index.js');

if (!existsSync(offlineNode)) {
  console.error(`[Error] Firewalled offline Node binary not found at: ${offlineNode}`);
  console.error(`Please run the following in an elevated PowerShell prompt (Run as Administrator):`);
  console.error(`  powershell -ExecutionPolicy Bypass -File scripts/lockdown-firewall.ps1\n`);
  process.exit(1);
}

if (!existsSync(sidecarEntry)) {
  console.error(`[Error] Compiled sidecar not found at: ${sidecarEntry}`);
  console.error(`Please run 'pnpm build' first.\n`);
  process.exit(1);
}

console.log(`[Excel AI] Launching sidecar with firewalled binary: ${offlineNode}`);
const child = spawn(offlineNode, ['--no-warnings', sidecarEntry], { stdio: 'inherit' });
child.on('exit', (code) => process.exit(code ?? 0));
