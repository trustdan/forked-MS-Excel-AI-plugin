# MS Excel AI Plugin

**The privacy-first, bring-your-own-model AI assistant for Microsoft Excel.** Open source, 100% local, air-gapped by design.

Connect **any** local or self-hosted AI model (LM Studio, Ollama, embedded in-process GGUF, or any OpenAI-compatible server) or private cloud endpoints (Claude, OpenAI) and let it analyze, query, and modify spreadsheets as an agent — with every change previewable and undoable before anything touches your grid.

---

## Key Highlights & Architectural Guarantees

* 🔒 **Air-Gap & Zero Telemetry:** Under no circumstances do workbook contents, cell formulas, user prompts, API keys, or operational telemetry leave `localhost`. An active network guard (`assertLoopbackUrl`) strictly blocks any non-loopback egress at the network socket layer.
* 🛡️ **Safe Agentic Edits with Diff Inspection:** Writes are staged into structured change-sets with user-visible justifications. A cell-level before/after diff inspector lets you review proposed changes (e.g. `Old: 100` → `New: =SUM(A1:A5)`) before clicking **Apply**, backed by automated formula regression testing and one-click **Undo**.
* ⚡ **Out-of-the-Box Local Default:** Pre-configured for local LM Studio (`http://127.0.0.1:1234/v1`) and Ollama (`http://127.0.0.1:11434/v1`), plus an **In-Process Embedded GGUF Engine** (`node-llama-cpp`) that runs local quantized models (Qwen 2.5, Llama 3.1) directly inside the sidecar with zero external servers.
* 🔌 **Model Context Protocol (MCP) Connectors:** Connect local SQLite databases, local CSV directories, and local filesystem trees to the agent, strictly sandboxed to explicit user-approved paths via `PathGuard` with prompt-injection defense.
* 📄 **Zero-Telemetry Local Document Importer:** Direct drag-and-drop or attachment of CSV, TSV, Markdown tables, and JSON arrays directly into worksheet change-sets with built-in formula injection escaping (CWE-1236).
* 🌓 **Adaptive Office Dark Theme:** Dynamically adapts to Excel's Dark Gray, Black, and Colorful themes (`Office.context.officeTheme` and OS `prefers-color-scheme`) with a 3-way toggle (Auto / Dark / Light).
* 🛑 **Instant Stop / Abort:** Full cancellation propagation (`AbortSignal`) across the agent loop, provider requests, custom functions (`CustomFunctions.Invocation.onCanceled`), and long-running MCP queries.
* 🎯 **Dynamic Range Interaction:** Live selection tracking ("Include Selection" chip) and interactive "Jump to Range" navigation that scrolls and highlights proposed change targets directly on the active worksheet.
* 🧪 **158 Automated Tests:** Exhaustive unit, component, integration, and E2E smoke tests ensuring zero regressions across all packages.

---

## Architecture Overview

```
 ┌─────────────────────────────────────────────────────────────┐
 │ Microsoft Excel (Desktop or Web)                            │
 │                                                             │
 │  ┌─────────────────────────┐   ┌─────────────────────────┐  │
 │  │ Office.js Taskpane UI   │   │ Custom Functions Runtime│  │
 │  │ (React + Vite + Dark UX)│   │ (=EXCELAI.AI, =CLASSIFY)│  │
 │  └────────────▲────────────┘   └────────────▲────────────┘  │
 └───────────────┼─────────────────────────────┼───────────────┘
                 │ WebSocket (Loopback)        │ HTTP (Loopback)
 ┌───────────────▼─────────────────────────────▼───────────────┐
 │ Local Sidecar Daemon (Node.js >= 22.13, localhost:8923)     │
 │                                                             │
 │  • Loopback Network Guard (blocks non-127.0.0.1 egress)     │
 │  • Agent Loop (Schema-validated tools & repair cycles)      │
 │  • Change-Set Manager (Diff inspector & rollback engine)    │
 │  • MCP Manager (SQLite, CSV directory, Filesystem sandbox)  │
 │  • Importers (CSV / TSV / Markdown / JSON parser)           │
 │  • Native SQLite Store (Token auth, cache, memory, ledger)  │
 │  • Embedded GGUF Engine (node-llama-cpp in-process LLM)     │
 └───────────────────────────────▲─────────────────────────────┘
                                 │ Loopback HTTP
 ┌───────────────────────────────▼─────────────────────────────┐
 │ Local Model Backends                                        │
 │ LM Studio (127.0.0.1:1234) · Ollama (127.0.0.1:11434)      │
 └─────────────────────────────────────────────────────────────┘
```

---

## Getting Started

### Prerequisites

* [Node.js](https://nodejs.org) $\ge$ 22.13.0
* `pnpm` $\ge$ 9.0 (`corepack enable`)
* Microsoft Excel on Windows (Desktop) or [Excel on the Web](https://office.com)
* A local model provider (e.g. [LM Studio](https://lmstudio.ai/) or [Ollama](https://ollama.com/)) or a local GGUF model file

---

### 1. Build and Run

Clone the repository and run the setup commands:

```powershell
# 1. Install dependencies without running untrusted third-party scripts
pnpm install --ignore-scripts

# 2. Verify type safety and test suite (158 tests)
pnpm check
pnpm test

# 3. Build production bundles
pnpm build

# 4. Start the sidecar companion daemon
pnpm dev:sidecar
```

When started, the sidecar logs:
```text
Excel AI sidecar listening on http://127.0.0.1:8923
Pairing code (enter it in the add-in once): AB12CD
Press Ctrl+C to stop.
```

> [!TIP]
> **Windows Defender Firewall Isolation:** To enforce binary-level outbound network sandboxing on Windows, run `powershell -ExecutionPolicy Bypass -File scripts/lockdown-firewall.ps1` to configure a dedicated offline Node binary and start via `pnpm start:offline`.

---

### 2. Connect from Microsoft Excel

No executable installer (`.exe`) is needed. Modern Office add-ins run securely inside Excel's native webview and communicate with the local sidecar over loopback.

#### Option A: Excel on the Web (Fastest — 10 seconds)
1. Open a blank workbook in [Excel for Web (office.com)](https://office.com).
2. On the ribbon, go to **Home** $\rightarrow$ **Add-ins** $\rightarrow$ **More Add-ins** (or **Insert** $\rightarrow$ **Office Add-ins**).
3. Select **Upload My Add-in** $\rightarrow$ browse to and select [`addin/manifest.xml`](addin/manifest.xml) $\rightarrow$ click **Upload**.

#### Option B: Excel Desktop on Windows
1. **Share the manifest folder:**
   - In File Explorer, right-click the `addin` folder $\rightarrow$ **Properties** $\rightarrow$ **Sharing** tab $\rightarrow$ click **Share...** $\rightarrow$ share with your Windows user and copy the network path (e.g. `\\YOUR-PC\addin`).
2. **Add to Excel Trusted Catalogs:**
   - In Excel: **File** $\rightarrow$ **Options** $\rightarrow$ **Trust Center** $\rightarrow$ **Trust Center Settings...** $\rightarrow$ **Trusted Add-in Catalogs**.
   - Paste the network path in **Catalog Url** $\rightarrow$ click **Add catalog** $\rightarrow$ check **Show in Menu** $\rightarrow$ click **OK**.
3. **Insert the Add-in:**
   - Restart Excel $\rightarrow$ open any workbook $\rightarrow$ **Insert** $\rightarrow$ **My Add-ins** $\rightarrow$ select the **Shared Folder** tab $\rightarrow$ choose **Excel AI** $\rightarrow$ click **Add**.

*(Alternatively, run `npx office-addin-debugging start addin/manifest.xml` if using the Office Add-in developer CLI).*

---

### 3. One-Time Pairing & Model Configuration

1. Click the **Excel AI** button on the Home ribbon to open the taskpane.
2. Enter the 6-character pairing code printed by the sidecar (e.g. `AB12CD`). Pairing is permanent and stored in `~/.excelai/settings.json` (mode `0600`).
3. In **Settings**:
   * **LM Studio (Default):** Ready out of the box at `http://127.0.0.1:1234/v1`.
   * **Ollama:** Click **Use Ollama defaults** (`http://127.0.0.1:11434/v1`) $\rightarrow$ click **Scan / List models** to pick from your installed local models.
   * **Embedded GGUF:** Select *Embedded GGUF* to run local quantized weights completely in-process using `node-llama-cpp`.
   * **MCP Connectors:** Switch to the **MCP Connectors** tab to whitelist approved folders and attach local SQLite databases or CSV collections.

---

## Features Walkthrough

### 💬 Conversational Spreadsheet Agent
* **Explore & Plan:** *"What sheets and tables are in this workbook?"* or *"Summarize the monthly trends in Sheet1!B2:G20."*
* **Formulas over Literals:** The agent prioritizes canonical en-US formulas (e.g. `=SUM(...)`, `=XLOOKUP(...)`) so your spreadsheet remains dynamic and maintainable.
* **Selection Attachment:** Click the **Selection: Sheet1!B2:F15** chip above the chat composer to insert the active Excel range into your prompt.

### 🔍 Transparent Diff Inspector & Change-Sets
* When the agent proposes workbook modifications, nothing changes immediately.
* An interactive **Change-Set Card** appears in the chat showing each affected cell with its prior value and proposed new value/formula.
* Click any range address to navigate directly to it on the active sheet.
* Click **Apply** to write changes to Excel with atomic verification, or **Undo** to roll back at any time.

### 🗄️ Model Context Protocol (MCP) Connectors
* Go to **Settings** $\rightarrow$ **MCP Connectors**.
* Add trusted directories under **Explicit User-Approved Paths** (e.g. `C:\Users\User\Documents\Data`).
* Configure connectors:
  * **SQLite Database:** Query local SQLite tables in read-only mode or stage table rows directly into a new sheet.
  * **CSV Directory:** Search, preview, and stage local CSV/TSV collections into worksheets.
  * **Filesystem:** Inspect directories and read text files safely within your approved sandbox.
  * **Custom Stdio Servers:** Connect external MCP servers over JSON-RPC 2.0 with full cancellation and lifecycle safety.

### 📎 Local Tabular File Importer
* Click the **📎 (Attach File)** icon in the chat composer to select a local `.csv`, `.tsv`, `.md`, or `.json` file.
* Preview table headers and dimensions directly in the composer chip.
* Ask the agent to import, format, and structure the data into any worksheet.

### ⚡ Batch AI Custom Functions
Use AI directly in Excel formulas (cached and batched to prevent UI freezes):
```excel
=EXCELAI.AI("Summarize in 5 words: " & A2)
=EXCELAI.AI_CLASSIFY(B2, "Support,Billing,Sales,Spam")
=EXCELAI.AI_EXTRACT(C2, "email address")
=EXCELAI.AI_TRANSLATE(D2, "Spanish")
```

---

## Workspace Layout

```
.
├── addin/               # Office.js React taskpane + custom functions
│   ├── src/taskpane/    # UI components (ChatBubble, Composer, ModelsScreen, McpScreen)
│   ├── src/officeBridge.ts # Office.js workbook execution engine & batch context.sync
│   ├── manifest.xml     # Office add-in manifest
│   └── test/            # Vitest + React Testing Library component tests
├── sidecar/             # Local Node.js daemon (Node >= 22.13 with native node:sqlite)
│   ├── src/agent/       # Agent loop, tool schemas, repair cycles, prompt defense
│   ├── src/mcp/         # MCP subsystem (PathGuard, SQLite, CSV, FS, stdio client)
│   ├── src/importers/   # Pure TypeScript tabular parsers (CSV, TSV, Markdown, JSON)
│   ├── src/providers/   # LM Studio / Ollama / OpenAI / Anthropic / Embedded GGUF
│   ├── src/server.ts    # Fastify-style HTTP + WebSocket server with pairing auth
│   ├── src/networkGuard.ts # Central loopback-only assertion guard
│   └── test/            # Vitest unit and integration test suites
├── shared/              # Shared TypeScript definitions, schemas, and A1 math
└── scripts/             # Windows Defender Firewall lockdown and offline scripts
```

---

## Testing & Verification

Run the full verification pipeline:

```powershell
# Type check all packages
pnpm check

# Run all 158 unit and integration tests
pnpm test

# Build production distribution
pnpm build

# Run end-to-end full-stack smoke test
pnpm smoke
```

---

## License

[MIT](LICENSE) © 2026 MS-Excel-AI-plugin contributors.
