# Excel AI Plugin — Developer & Agent Guide

Open-source, strictly local-first, air-gapped AI plugin for Microsoft Excel.
Architecture: thin Office.js add-in (`addin/`) + local Node.js sidecar (`sidecar/`)
sharing protocol types and tool schemas (`shared/`). Design doc: `docs/DESIGN.md`.

---

## Core Commands

| Command | What it does |
|---|---|
| `pnpm install --ignore-scripts` | Install workspace dependencies safely without running lifecycle scripts |
| `pnpm check` | Typecheck every package (`shared` built first) |
| `pnpm test` | Vitest unit + integration suites (runs against the headless workbook emulator) |
| `pnpm build` | Build workspace: `shared` → `sidecar` → `addin` |
| `pnpm smoke` | Full-stack E2E test: live sidecar + mock OpenAI-compatible model + fake add-in over WS |
| `pnpm dev:sidecar` | Build & run the sidecar in development mode |
| `pnpm start:offline` | Run sidecar with the firewalled offline Node binary (`./bin/node-offline.exe`) |

---

## Non-Negotiable Rules & Invariants

1. **Air-Gap & Zero Exfiltration:**
   - Outbound network traffic is strictly restricted to loopback (`127.0.0.1`, `localhost`, `[::1]`).
   - Every outbound model request is validated through `assertLoopbackUrl`. Non-local destinations throw `NonLocalConnectionBlockedError`.
   - Never introduce cloud provider fallback URLs (`api.openai.com`, `api.anthropic.com`).
   - Default provider is OpenAI-compatible pointed at local LM Studio (`http://127.0.0.1:1234/v1`).
2. **Quality Gate:**
   - A change is only done when `pnpm check && pnpm test && pnpm build && pnpm smoke` all pass.
3. **Single Source of Truth for Tools:**
   - `shared/src/tools.ts` defines all workbook tool schemas. If you modify it, update BOTH implementations:
     - `sidecar/src/workbook/emulator.ts` (headless emulator)
     - `addin/src/officeBridge.ts` (Office.js bridge)
     - plus the corresponding agent test suites.
4. **Formula Canonicalization:**
   - All formulas in the protocol are canonical en-US (`formulas`, never `formulasLocal`). Use standard en-US function names and comma separators.
5. **Change-Set Atomicity:**
   - Write tools must stage into change-sets — never apply directly. Every write op requires a descriptive `reason`.
6. **Data Isolation:**
   - Tests must never touch `~/.excelai` (vitest setup redirects `EXCELAI_DATA_DIR` to a temporary directory).
7. **Formula Parity:**
   - The emulator's formula engine must implement functions exactly or return `#NAME?` — never approximate results.
8. **Dependency Constraints:**
   - Zero GPL dependencies. No lifecycle scripts (`--ignore-scripts`).
9. **Multi-Agent Collaboration:**
   - Adhere to the specialist personas defined in `AGENTS.md` (`bdd-expert`, `devsecops-engineer`, `llm-engineer`, `ms-excel-specialist`, `cyber-security-professional`, `structural-integrity-manager`).

---

## Repository Layout

- `shared/src/` — Tool schemas (`zod`), WS protocol envelopes, change-set structures, A1 notation utilities
- `sidecar/src/agent/` — Tool-calling loop, schema validation, repair-retry, system prompts
- `sidecar/src/changeset/` — Stage → preview → approve → apply → verify (auto-rollback) → undo
- `sidecar/src/providers/` — OpenAI-compatible (LM Studio / Ollama / vLLM), Anthropic, scripted (tests)
- `sidecar/src/networkGuard.ts` — Loopback network guard blocking all non-local connections
- `sidecar/src/workbook/` — Workbook executor interface, headless emulator, formula engine
- `sidecar/test/` — Unit and integration tests; `sidecar/scripts/smoke.mjs` — E2E smoke test
- `addin/src/` — React taskpane UI, Office.js bridge, custom functions (`=EXCELAI.AI()` family)
- `scripts/lockdown-firewall.ps1` — Windows Defender Firewall isolation script for offline Node execution

---

## Testing Against Real Excel (Manual)

1. Start your local model in LM Studio (ensure the local server is running on `http://127.0.0.1:1234`).
2. Build and start the sidecar:
   ```powershell
   pnpm build && pnpm dev:sidecar
   ```
   (Serves the taskpane at `http://127.0.0.1:8923`)
3. Sideload `addin/manifest.xml` into Excel (Web: Insert → Add-ins → Upload My Add-in; or Windows desktop add-in catalog).
4. Enter the one-time pairing code printed in the sidecar console into the taskpane.
5. In Settings, the default LM Studio configuration is pre-selected and ready to test.
