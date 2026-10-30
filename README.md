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
* 🧪 **162 Automated Tests:** Exhaustive unit, component, integration, and E2E smoke tests ensuring zero regressions across all packages.

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

# 2. Verify type safety and test suite (162 tests)
pnpm check
pnpm test

# 3. Build production bundles
pnpm build

# 4. Start the sidecar companion daemon (or use `pnpm dev` / `pnpm start`)
pnpm dev:sidecar
```

When started, the sidecar logs:
```text
Excel AI sidecar listening on http://127.0.0.1:8923
Pairing code (enter it in the add-in once): AB12CD
Press Ctrl+C to stop.
```

> [!TIP]
> **Windows Defender Firewall Isolation:** To enforce binary-level outbound network sandboxing on Windows, run `powershell -ExecutionPolicy Bypass -File scripts/lockdown-firewall.ps1` to configure a dedicated offline Node binary and start via `pnpm start:offline`. The rule blocks every non-loopback IPv4 and IPv6 address for that one executable. Stdio MCP servers run as separate processes: unless a server's command is `bin/node-offline.exe`, it is not covered and needs its own outbound block rule.

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

### 4. Hardware Tiers & Recommended Local Model Settings

Because the agent sends workbook tool schemas, multi-turn conversation history, and spreadsheet range data with each turn—and **reasoning/thinking models** also spend output tokens on internal `<think>` / `reasoning_content` blocks before emitting a tool call—default **4K or 8K (`8192`) context windows** in LM Studio or Ollama can easily run out of tokens (`finish_reason: "length"` / *"The model ran out of output tokens before producing a reply"*).

Use the table below to tune **LM Studio**, **Ollama**, or **Embedded GGUF** for your machine's memory:

| System RAM / Unified Memory | Recommended Model Size & Quant | Context Length (`n_ctx`) | KV Cache & Attention Settings | Best For |
| :--- | :--- | :--- | :--- | :--- |
| **16 GB RAM** *(Entry / Laptop)* | **7B – 8B** (`Q4_K_M`)<br>*(e.g. Qwen 2.5 7B, Llama 3.1 8B)* | **`16384` (16K)**<br>*(8K minimum; avoid reasoning models on 8K)* | • **Flash Attention:** ON<br>• **KV Cache:** `Q8_0` (or `Q4_0` if VRAM-tight) | Single-sheet edits, formula generation, and targeted range queries (`A1:G30`). |
| **32 GB – 64 GB RAM** *(Workstation / Gaming PC)* | **14B – 32B** (`Q4_K_M` – `Q6_K`)<br>*(e.g. Qwen 2.5 14B/32B, DeepSeek-R1 Distill)* | **`32768` (32K)** | • **Flash Attention:** ON<br>• **KV Cache:** `Q8_0` or `FP16` | Multi-sheet analysis, reasoning models, and MCP database/CSV queries. |
| **128 GB – 192 GB+ RAM** *(High-Memory Workstation)* | **32B – 70B+ / MoE** (`Q6_K`, `Q8_0`, or `FP16`) | **`65536` (64K) – `131072` (128K)** | • **Flash Attention:** ON<br>• **KV Cache:** Unquantized `FP16` (or `Q8_0`) | Long multi-turn sessions, deep reasoning models, and large tabular imports with maximum numerical precision. |

#### Practical Tuning Tips

* **Why 64K–128K on High-Memory Machines (128 GB – 192 GB+ RAM):**
  * Setting the context slider in LM Studio to **`65536` (64K)** or **`131072` (128K)** only pre-allocates KV cache RAM (which you have plenty of)—prompt processing speed only scales with how many tokens are *actually* in the current conversation.
  * Keep **KV Cache Quantization** at **unquantized `FP16`** (or `Q8_0`) rather than `Q4_0`: aggressive KV cache quantization can degrade precision on exact cell coordinates (`Sheet1!D14:G42`) and numerical values.
  * Always enable **Flash Attention** in LM Studio's Model Loader settings to accelerate prompt evaluation and reduce KV cache overhead.
* **Using Reasoning / "Thinking" Models:**
  * Reasoning models generate internal chain-of-thought tokens (`reasoning_content`) before emitting tool calls or replies. Ensure your context window is at least **16K–32K+** so a 6K–7K prompt doesn't leave the model with only ~1,000 tokens before hitting an 8K ceiling.
  * Prefer **specific requests** (e.g., *"Build a 5-row monthly budget table in Sheet1!A1:D6 with SUM formulas"*) over open-ended prompts (*"Can you build something and show me?"*), which can cause reasoning models to spend 1,000+ tokens deliberating what sample topic to invent.
* **Ollama Context Override (`num_ctx`):**
  * Ollama defaults to a `2048` or `4096` token context window unless overridden. To run with a 32K or 64K window in Ollama, create a `Modelfile` with `PARAMETER num_ctx 32768` (or `65536`) or set `OLLAMA_CONTEXT_LENGTH=32768` in your environment before starting `ollama serve`.
* **Sidecar Power-User Knobs & Loop Safeguards ([`sidecar/src/agent/loop.ts`](sidecar/src/agent/loop.ts)):**
  * `maxTokens: 16384` caps the maximum completion tokens (reasoning + tool calls) per turn. On 64K–128K context setups with heavy reasoning models, you can raise this to `32768`.
  * `maxIterations = 24` sets the maximum tool-calling steps per turn. On the final step (or immediately if a local model repeats an identical tool call), the sidecar automatically disables `tools` for one synthesis step so the model summarizes its progress or asks a clarifying question in plain text instead of timing out.
  * **Reasoning Exhaustion Recovery:** If a local reasoning model (`reasoning_content`) exhausts its token budget mid-thought without emitting a tool call or reply, the agent loop automatically detects it (even when the server reports `finish_reason: "stop"`) and sends one continuation nudge to emit the tool call or answer directly.
  * `MAX_TOOL_RESULT_CHARS = 20_000` caps the size of a single `read_range` call (~5,000 tokens) to protect smaller context windows. On 64K–128K setups, you can increase this if you want the agent to inspect larger cell ranges in a single step.

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
* Add trusted directories under **Explicit User-Approved Paths** (e.g. `C:\Users\User\Documents\Data`). These folders are also the only places the agent's file import may read from; to import a file outside them, attach it in chat instead.
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

Each saved model configuration has its own cache, so switching models never returns another model's answers. Changing a configuration's endpoint, model or model file starts a fresh cache; changing only its API key does not. Cached results expire after 30 days, so a model replaced behind the same name and endpoint is eventually picked up. If more than 256 calls are waiting for the model, extra calls return an error; recalculate once the backlog clears.

### 🗂️ Local data and retention

Everything lives in `~/.excelai` (or `EXCELAI_DATA_DIR`), and nothing grows without bound:

| Data | Kept |
|---|---|
| `=EXCELAI.AI()` result cache | newest 20,000 results, up to 30 days |
| Usage ledger (token counts) | newest 200,000 calls, up to 400 days; totals cover this window |
| Change-sets (for review and undo) | per workbook session: newest 200, up to 24 hours, at most 1,000,000 snapshot cells; older ones can no longer be undone |

`settings.json` is written atomically (temp file, fsync, rename). If it cannot be parsed, or has the wrong shape, it is moved to `settings.json.corrupt-<timestamp>`, the sidecar starts with empty settings, and the Models screen says where the original went. If the file exists but cannot be read at all (for example, it is locked), the sidecar refuses to touch it and settings requests fail until it is readable again.

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

# Run all 162 unit and integration tests
pnpm test

# Build production distribution
pnpm build

# Run end-to-end full-stack smoke test
pnpm smoke
```

---

## License

[MIT](LICENSE) © 2026 MS-Excel-AI-plugin contributors.
