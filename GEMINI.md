# Gemini & Antigravity Agent Guidelines — Excel AI Plugin

This repository houses an open-source, air-gapped, bring-your-own-model AI plugin for Microsoft Excel.
The architecture consists of a thin Office.js add-in (`addin/`) paired with a local Node.js sidecar daemon (`sidecar/`) communicating over loopback WebSockets and HTTP.

---

## Prime Directive: Air-Gap & Zero Telemetry

Under no circumstances may workbook contents, cell formulas, user prompts, API keys, or operational telemetry leave `localhost`.
* **Allowed Hostnames:** `127.0.0.1`, `localhost`, `[::1]`.
* **Prohibited:** Any public or external IP addresses, cloud endpoints (e.g. `api.openai.com`, `api.anthropic.com`), or analytics services.
* **Network Enforcement:** Outbound requests in `sidecar/src/providers/` and model discovery endpoints must pass through `assertLoopbackUrl()`. Any violation must immediately throw `NonLocalConnectionBlockedError`.
* **Default Model Backend:** Local LM Studio server running at `http://127.0.0.1:1234/v1`.

---

## Agent Personas & Operational Roles

When tackling any task in this repo, align with the personas codified in [AGENTS.md](file:///c:/Users/Dan/forked-MS-Excel-AI-plugin/AGENTS.md):
- **`bdd-expert`**: Define clear acceptance criteria and verify user-observable behavior.
- **`devsecops-engineer`**: Safeguard supply chain, lockfile integrity, `--ignore-scripts`, and firewall policies.
- **`llm-engineer`**: Ensure local quantized models (Llama 3, Qwen) work with structured JSON tools and prompt budgets.
- **`ms-excel-specialist`**: Guard formula parity (en-US canonical), Office.js bridge, and headless emulator synchronization.
- **`cyber-security-professional`**: Verify loopback guards, token authentication, SQLite FTS safety, and input sanitization.
- **`structural-integrity-manager`**: Ensure memory-bounded buffers, cancellation propagation, and rollback safety.

---

## Workspace Layout & Key Files

* [shared/](file:///c:/Users/Dan/forked-MS-Excel-AI-plugin/shared) — Shared TypeScript types, zod schemas, protocol envelopes, A1 address math.
  * [shared/src/tools.ts](file:///c:/Users/Dan/forked-MS-Excel-AI-plugin/shared/src/tools.ts) — The single source of truth for the workbook tool surface.
* [sidecar/](file:///c:/Users/Dan/forked-MS-Excel-AI-plugin/sidecar) — Local Node.js background process (runs on Node >= 22.13 using native `node:sqlite`).
  * [sidecar/src/networkGuard.ts](file:///c:/Users/Dan/forked-MS-Excel-AI-plugin/sidecar/src/networkGuard.ts) — Central loopback URL validator.
  * [sidecar/src/server.ts](file:///c:/Users/Dan/forked-MS-Excel-AI-plugin/sidecar/src/server.ts) — HTTP + WebSocket server, token authentication, static asset hosting.
  * [sidecar/src/config.ts](file:///c:/Users/Dan/forked-MS-Excel-AI-plugin/sidecar/src/config.ts) — Local configuration manager (`~/.excelai/settings.json`, mode `0600`).
  * [sidecar/src/providers/](file:///c:/Users/Dan/forked-MS-Excel-AI-plugin/sidecar/src/providers) — Model provider adapters (OpenAI-compatible / LM Studio / Ollama, Anthropic).
  * [sidecar/src/workbook/](file:///c:/Users/Dan/forked-MS-Excel-AI-plugin/sidecar/src/workbook) — Headless emulator and formula evaluation engine.
  * [sidecar/test/](file:///c:/Users/Dan/forked-MS-Excel-AI-plugin/sidecar/test) — Vitest suites.
* [addin/](file:///c:/Users/Dan/forked-MS-Excel-AI-plugin/addin) — Office.js React taskpane UI + Custom Functions.
  * [addin/src/taskpane/App.tsx](file:///c:/Users/Dan/forked-MS-Excel-AI-plugin/addin/src/taskpane/App.tsx) — Main taskpane UI and provider settings.
  * [addin/src/officeBridge.ts](file:///c:/Users/Dan/forked-MS-Excel-AI-plugin/addin/src/officeBridge.ts) — Office.js execution bridge for tool calls.
  * [addin/manifest.xml](file:///c:/Users/Dan/forked-MS-Excel-AI-plugin/addin/manifest.xml) — Excel add-in manifest.
* [scripts/lockdown-firewall.ps1](file:///c:/Users/Dan/forked-MS-Excel-AI-plugin/scripts/lockdown-firewall.ps1) — PowerShell script for Windows Defender Firewall lockdown.

---

## Standard Development & Verification Workflow

```powershell
# 1. Install dependencies without running untrusted scripts
pnpm install --ignore-scripts

# 2. Typecheck all packages
pnpm check

# 3. Run unit and integration tests
pnpm test

# 4. Build all packages
pnpm build

# 5. Full end-to-end smoke test
pnpm smoke
```

Every PR, refactor, or hardening change must pass `pnpm check && pnpm test && pnpm build && pnpm smoke`.
