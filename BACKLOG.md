# Product & Engineering Backlog — Excel AI Plugin

This backlog tracks the architectural, security, testing, and feature tasks for the `excel-ai-plugin` project.

---

## Status Legend
- 🟢 **Done** — Implemented, tested, and verified.
- 🟡 **In Progress** — Currently under active development.
- ⚪ **Backlog** — Scoped and prioritized for upcoming iterations.

---

## Epic 1: Air-Gap Hardening & Offline LM Studio Default
**Goal:** Guarantee zero workbook data, prompts, or telemetry leave `localhost`. Enforce strict loopback-only communication and provide out-of-the-box LM Studio integration.

- [x] **SEC-01 (P0): Loopback Network Guard** 🟢
  - Implement `assertLoopbackUrl` in `sidecar/src/networkGuard.ts` strictly restricting outbound requests to `127.0.0.1`, `localhost`, or `[::1]`.
  - Throw `NonLocalConnectionBlockedError` before any socket or fetch is created.
  - Integrate guard into `OpenAIProvider`, `AnthropicProvider`, and `listModels` discovery endpoint.
- [x] **SEC-02 (P0): Strip Cloud Endpoints & Default Fallbacks** 🟢
  - Remove all hardcoded references to `https://api.openai.com/v1` and `https://api.anthropic.com`.
  - Prevent accidental fallback to public clouds if a configuration value is missing.
- [x] **SEC-03 (P0): Out-of-the-Box LM Studio Default** 🟢
  - Seed default configuration in `sidecar/src/config.ts` targeting `http://127.0.0.1:1234/v1` with model `local-model` and fallback key `lm-studio`.
  - Update `addin/src/taskpane/App.tsx` default form to OpenAI-compatible / LM Studio.
  - Add quick preset buttons in the taskpane for LM Studio (`127.0.0.1:1234`) and Ollama (`127.0.0.1:11434`).
- [x] **SEC-04 (P1): Windows Defender Firewall Lockdown Script** 🟢
  - Provide `scripts/lockdown-firewall.ps1` to isolate a dedicated `node-offline.exe` binary.
  - Create Windows Defender outbound rule blocking non-loopback IP ranges (`0.0.0.0-126.255.255.255` and `128.0.0.0-255.255.255.255`).
  - Add `start:offline` scripts in `package.json`.
- [x] **SEC-05 (P0): Safe Verification Suite & Network Guard Tests** 🟢
  - Add unit test suite `sidecar/test/network-guard.test.ts` proving all external IPs and domain names are rejected prior to network transmission.
  - Run verification across `pnpm check`, `pnpm test`, and `pnpm smoke`.

---

## Epic 2: Multi-Agent Architecture & Specialist Personas
**Goal:** Define and operationalize domain-specialist personas in `AGENTS.md` to guide ongoing development, code reviews, and automated verification.

- [x] **AGT-01 (P0): Create `AGENTS.md` Specialist Charters** 🟢
  - `bdd-expert`: User story framing, Gherkin specs, end-to-end Excel workflows, regression acceptance criteria.
  - `devsecops-engineer`: Air-gap lockdown, CI/CD pipeline integrity, firewall policies, `--ignore-scripts` enforcement, SBOM & dependency auditing.
  - `llm-engineer`: Local LM Studio / Ollama model tuning, GGUF schema-constrained generation, context window optimization, tool-call repair loops.
  - `ms-excel-specialist`: Office.js API mastery, Shared Runtime lifecycle, custom functions (`=EXCELAI.AI`), canonical formula syntax, VBA/Python interoperability.
  - `cyber-security-professional`: Threat modeling, zero-exfiltration audits, token-based pairing auth, CORS origin enforcement, SQLite query sanitization.
  - `structural-integrity-manager`: Memory management, zero-leak streaming, process lifecycle, cancellation token propagation, cache eviction, crash resilience.

---

## Epic 3: AI Development & Coding Standard Docs
**Goal:** Align developer guidelines and AI assistant instructions across tooling environments.

- [x] **DOC-01 (P0): Enhance `CLAUDE.md`** 🟢
  - Include offline LM Studio operational rules, network guard invariants, and testing mandates.
- [x] **DOC-02 (P0): Create `GEMINI.md`** 🟢
  - Provide comprehensive Antigravity / Gemini coding and architecture guide for the repository.

---

## Epic 4: Advanced Excel Agent Features & Model Connectivity (Future)
**Goal:** Deepen Excel analysis capabilities while retaining 100% local privacy.

- [x] **FEAT-01 (P2): Embedded In-Process LLM Engine (`node-llama-cpp`)** 🟢
  - Implemented `EmbeddedLlamaProvider` (`sidecar/src/providers/embedded.ts`) with dynamic optional `node-llama-cpp` loading and deterministic `MockLlamaBackend` for automated test suites.
  - Added secure model resolution with path traversal guards (`resolveModelPath`) and automatic local GGUF discovery (`listLocalGgufModels`) in `~/.excelai/models/`.
  - Added prompt formatting for local quantized models (Qwen 2.5 / Llama 3.1) with JSON schema tool definitions and regex/markdown tool-call extraction (`extractToolCalls`).
  - Added AbortSignal cancellation and thread/context size bounding for in-process execution.
  - Added "Embedded GGUF" selector in taskpane settings with "Scan local models" trigger and full unit test coverage.
- [x] **FEAT-02 (P2): Local Document Importer (CSV / TSV / Markdown / JSON)** 🟢
  - Implemented pure TypeScript zero-telemetry tabular parser (`sidecar/src/importers/tabular.ts`) supporting RFC 4180 CSV, TSV, Markdown tables, and JSON arrays.
  - Added formula injection defense (CWE-1236) escaping leading formula triggers (`=`, `@`, `+`, `-`) in untrusted text cells while preserving valid signed numbers.
  - Added `import_local_data` tool in `shared/src/tools.ts` and `sidecar/src/agent/loop.ts` staging `write_range` change-sets with optional header formatting and preview.
  - Added `/api/importer/parse` endpoint for local file inspection and table preview.
  - Added file attachment picker (📎) and file chip in `addin/src/taskpane/components/chat/Composer.tsx` with full unit test coverage.
- [x] **FEAT-03 (P3): Model Context Protocol (MCP) Local Connectors** 🟢
  - Implemented strict sandboxing validator `PathGuard` (`sidecar/src/mcp/pathGuard.ts`) restricting all filesystem, database, and CSV access to explicit user-approved paths, with path canonicalization and directory traversal prevention.
  - Implemented built-in read-only SQLite connector (`sidecar/src/mcp/connectors/sqliteConnector.ts`) with `sqlite_list_tables`, `sqlite_describe_table`, read-only `sqlite_read_query` (blocking mutation statements), and `sqlite_stage_table_to_sheet` for staging queries directly into worksheet change-sets.
  - Implemented built-in CSV Directory connector (`sidecar/src/mcp/connectors/csvDirectoryConnector.ts`) with `csv_list_files`, `csv_read_preview`, `csv_query`, and `csv_stage_to_sheet`.
  - Implemented built-in Filesystem connector (`sidecar/src/mcp/connectors/filesystemConnector.ts`) with `fs_list_directory`, `fs_read_file` (with binary rejection), `fs_search_files`, and `fs_get_metadata`.
  - Implemented standard JSON-RPC 2.0 stdio MCP client (`sidecar/src/mcp/transports/stdioTransport.ts`) supporting arbitrary external local MCP servers with protocol handshake, tool discovery, execution, lifecycle cleanup, and AbortSignal cancellation.
  - Built centralized `McpManager` (`sidecar/src/mcp/manager.ts`) coordinating server connections, dynamic tool spec exposure to LLMs, prompt-injection defense with untrusted data wrappers, and testing.
  - Added HTTP management endpoints (`/api/mcp/paths`, `/api/mcp/servers`, `/api/mcp/servers/test`, `/api/mcp/tools`) and taskpane settings UI tab (`addin/src/taskpane/components/settings/McpScreen.tsx`) for managing approved paths and connectors.
  - Added comprehensive test suites across path guarding, SQLite, CSV, filesystem, stdio protocol, manager, API endpoints, agent loop turns, and React UI components.

---

## Epic 5: Add-In Modernization, Rich UX & Office.js Deepening
**Goal:** Transform the Office.js add-in from a minimal prototype into a responsive, accessible, dark-mode aware, and visual spreadsheet assistant with cell-level diffs and test coverage.

- [x] **UI-01 (P0): Rich Markdown & Code Block Formatting** 🟢
  - Render assistant messages with structured formatting (bold, italic, numbered lists, tables, inline code blocks) using a zero-dependency or lightweight local parser.
  - Eliminate unformatted raw text dumps in the chat pane.
- [x] **UI-02 (P0): Visual Before / After Diff Inspector for Change-Sets** 🟢
  - Expand change-set op rows to display cell-level before-and-after values (e.g. `Old: 100` → `New: =SUM(A1:A5)`).
  - Give users transparent visibility into exactly what cells will change before clicking "Apply".
- [x] **UI-03 (P1): Office Dark Theme Adaptation** 🟢
  - Listen to `Office.context.officeTheme` and `prefers-color-scheme`, dynamically toggle taskpane CSS custom properties between light and dark modes with persistent 3-way toggle (Auto / Dark / Light).
  - Provide full visual harmony with Excel's Dark Gray and Black themes, including specialized contrast styling for diff inspectors, markdown tables, and status badges.
- [x] **UI-04 (P1): Stop / Abort In-Flight Generation** 🟢
  - Add an interrupt/stop button in the chat composer to cancel running agent loops immediately via WebSocket `abort` signal.
  - Wire `AbortController` and `AbortSignal` through `Session`, `runAgentTurn()`, and provider `fetch` calls with fast-fail retry bypassing.
  - Gracefully abort staged changes, emit `stopped` events, prevent history corruption, and provide keyboard shortcut (Escape).
- [x] **BRG-01 (P0): "Include Selection" Context Chip** 🟢
  - Add a quick-attach chip above the chat composer displaying the currently selected Excel range (e.g., `Selection: Sheet1!B2:F15`).
  - Allow one-click insertion of the active selection into the user's prompt.
- [x] **BRG-02 (P1): Interactive "Jump to Range" Navigation** 🟢
  - Make range addresses in change-set previews clickable to trigger `range.select()`.
  - Automatically scroll and highlight the target area on the active sheet so users can inspect proposed targets.
- [x] **BRG-03 (P1): Batch `context.sync()` Performance Optimization** 🟢
  - Refactor `applyOps()` in `addin/src/officeBridge.ts` to eliminate sequential `await context.sync()` inside loop iterations.
  - Group independent cell mutations into chunks of 50 per sync, sync only at structural boundaries (sheets, row/col shifts), and use synchronous A1 dimension parsing for `numberFormat` matrices to eliminate redundant read syncs.
- [x] **CF-01 (P1): Invocation Cancellation for Custom Functions** 🟢
  - Connect `CustomFunctions.Invocation.onCanceled` in `addin/src/functions/functions.ts` to abort underlying sidecar requests when cell calculation is canceled.
  - Added `"options": { "cancelable": true }` to custom functions metadata, wired `AbortController` cancellation through HTTP `/api/functions/run`, and unblocked queued `FunctionEngine` slots safely.
- [x] **TST-01 (P1): Add-In Component Architecture & Unit Test Suite** 🟢
  - Decomposed monolithic `App.tsx` (formerly 861 lines) into modular components: `components/chat/` (`ChatBubble`, `Composer`), `components/changeset/` (`ChangeSetCard`), `components/settings/` (`ModelsScreen`), `components/pairing/` (`PairingScreen`), and shared types in `types.ts`.
  - Configured Vitest + `@testing-library/react` + `jsdom` for `addin/` with component unit tests covering pairing submission, message streaming, composer interactions (Enter, Shift+Enter, selection chip attachment, abort/stop button), and change-set review (diff inspector, approve/reject).

