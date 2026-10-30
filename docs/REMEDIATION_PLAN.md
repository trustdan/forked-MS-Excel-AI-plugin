# Remediation plan — pragmatic audit 2026-10-29

Source: [PRAGMATIC_AUDIT.md](PRAGMATIC_AUDIT.md) (reviewed commit `ea2f25d`). This plan groups every finding into epics. Each epic covers one boundary, and the epics are ordered into phases that can each be merged on their own.

References: **F1–F12** are the audit's numbered findings. **C1–C7** are rows of its "Disposition of the supplied review" table, in order. **L1–L4** are its "Additional lifecycle work" bullets, in order.

## Ground rules for every epic

- **Done means green:** `pnpm check && pnpm test && pnpm build && pnpm smoke`, and each finding closed in that epic has a regression test that fails on `ea2f25d`.
- **Test the real boundary:** test the production code path. Use real loopback HTTP/WS listeners, both workbook executors, and real files in the temp `EXCELAI_DATA_DIR`. Don't stub out the part being fixed (the audit's recurring complaint).
- **Executor parity:** any change to `shared/src/tools.ts` or `shared/src/protocol.ts` updates `sidecar/src/workbook/emulator.ts`, `addin/src/officeBridge.ts`, and both test suites in the same PR.
- **The audit's Acceptance lines are the spec.** Below, epics list only the extra criteria that are not already in the audit.
- **One PR per epic, or per story for large epics.** Never mix phases in one PR.

## Epics

| Epic | Title | Covers | Lead persona | Supporting |
| --- | --- | --- | --- | --- |
| E0 | Stop-gap hotfixes | F5 (crash), F6 (redirect), C7 (CI) | `cyber-security-professional` | `devsecops-engineer` |
| E1 | Exact workbook targeting | F1 | `ms-excel-specialist` | `bdd-expert` |
| E2 | Change-set state machine | F3, C1, F8 | `structural-integrity-manager` | `bdd-expert` |
| E3 | Formula meaning & address-level verification | F2, F4 | `ms-excel-specialist` | `llm-engineer` |
| E4 | Validated inbound protocol | F5 (full) | `cyber-security-professional` | `structural-integrity-manager` |
| E5 | One guarded outbound HTTP transport | F6 (full), C2, F7, C3 | `cyber-security-professional` | `llm-engineer` |
| E6 | Data-access capabilities & bounded work | F9, F10, F11, C4 | `cyber-security-professional` | `structural-integrity-manager` |
| E7 | Cache identity, persistence & retention | F12, C5, L3 | `structural-integrity-manager` | `llm-engineer` |
| E8 | Native & child-process lifecycle | L1, L2, L4 | `structural-integrity-manager` | `devsecops-engineer` |
| E9 | Docs/policy consistency | C6, C7 (README count) | `devsecops-engineer` | `bdd-expert` |

### E0 — Stop-gap hotfixes

Small, low-risk patches that close the demonstrated P1 crash and bypass while E4/E5 build the real fix.

- Wrap the detached WS message handler in `server.ts` with a `.catch`, reject non-object frames before reading `msg.v`, attach per-socket `error` handlers, and set an explicit `maxPayload`.
- Set `redirect: 'error'` on the model-discovery `fetch()` in `server.ts` (around lines 714–728).
- Add `--ignore-scripts` to the install step in `.github/workflows/ci.yml`.

**Acceptance:** sending a `null` frame closes only that socket; a discovery redirect fails; CI installs with `--frozen-lockfile --ignore-scripts`.

### E1 — Exact workbook targeting (F1)

- Remove the active-sheet fallback in `officeBridge.ts` `getWorksheet()` for every mutation and snapshot path. A missing sheet throws the same error the emulator throws.
- Resolve sheet names once, while the proposal is built. Store the resolved target in the change-set and show it in the preview.
- Track worksheet identity (Office.js worksheet `id`; a stable synthetic id in the emulator) so a rename between preview and apply is detected rather than guessed at. Allow case-insensitive name matching.
- Replace the add-in test that expects the `Sheet1` fallback with tests for the new behavior.

**Extra acceptance:** the preview shows the resolved sheet; renaming a target sheet after preview either follows the id or fails explicitly, never silently retargets.

### E2 — Change-set state machine (F3, C1, F8)

Large epic. Ship it as three stories, in order:

1. **E2a: Serialization.** Add a per-workbook async mutex, acquired before the first `await` in `apply`/`undo`/`reject`. Add explicit states: `staged → applying → applied | failed_partial`, then `applied → undoing → undone`. A duplicate approval returns the in-flight result or an `already_in_progress` response.
2. **E2b: Partial-apply recovery (C1).** Catch failures mid-apply, attempt rollback from the snapshot, and end in `rolled_back` or `failed_partial` with the list of unrecovered cells. Never leave a set `staged` after it has touched the workbook.
3. **E2c: Honest, conflict-aware undo (F8).**
   - Carry `undoFidelity` in the staged event and render it on `ChangeSetCard`.
   - Make add-sheet undo delete the sheet.
   - Snapshot formatting for sorts, or downgrade them to `partial`.
   - Before undoing, compare current cells with the set's applied result; on mismatch, report a conflict instead of overwriting. Allow only the newest applied set on a workbook to be undone, unless the comparison is clean.

**Extra acceptance:** the UI never shows Undo for a set that can't be undone cleanly; the new states survive a sidecar restart, or are reported as `unknown` after one.

### E3 — Formula meaning & address-level verification (F2, F4)

- **F2:** In `loop.ts`, stop `repairWriteRangeCells()` from rewriting references that already parse and resolve. Limit automatic repair to unambiguous syntax fixes, such as a missing `=`, a stray locale separator, or unbalanced parentheses. Send ambiguous references back to the model as a structured diagnostic (`{cell, formula, issue, candidates}`). Remove the unproven "all formulas valid" claim from the staging response (around line 684).
- **F4:** In `manager.ts`, replace the `beforeErrors === 0` gate with a per-address diff: a cell is a regression when its error identity is new or changed. Extend the cell protocol with real error metadata (both executors), so text such as `#SKU` is not an error. A failed verification re-read returns `verification_incomplete`, never `ok: true`.
- Document the verification scope (written ranges only, or written ranges plus direct dependents) in `docs/DESIGN.md`.

**Extra acceptance:** the audit's `Prior subtotal` probe and its `=1/0` / `=2/0` probe are kept as named regression tests.

### E4 — Validated inbound protocol (F5)

- Define zod schemas in `shared/src/protocol.ts` for every WS envelope and payload. Parse before dispatch; the TypeScript cast goes away.
- Reply to malformed frames with a typed protocol error, and close the socket after repeated violations.

**Depends on:** E0. **Extra acceptance:** a second client and `/health` stay responsive while one client sends garbage.

### E5 — One guarded outbound HTTP transport (F6, C2, F7, C3)

Replace the per-call-site `fetch()` and `loopbackFetch()` usage with a single transport in `networkGuard.ts`:

- Resolve the hostname, check the address is loopback, then connect to that exact address (C2). Redirects are rejected, or re-validated and re-pinned at every hop (F6). This applies to discovery, probes, chat, and custom functions.
- Return a streaming response at header time. Enforce byte and inactivity limits on the socket, and pass abort through to the socket (F7).
- Remove test-detection branches from the guard. Tests inject a transport or use local fixtures instead.
- **C3:** fix the Anthropic adapter's double `/v1`, and assert exact URLs in tests.

**Extra acceptance:** a live-HTTP test shows the first SSE delta reaching the caller before the server ends the stream; an IPv4-mapped IPv6 redirect is rejected.

### E6 — Data-access capabilities & bounded work (F9, F10, F11, C4)

- **F9:** The importer reads only user-attached content or paths under approved roots (`PathGuard`). The HTTP parser follows the same rule. If unrestricted import is kept on purpose, record that as an explicit policy exception in `docs/DESIGN.md`.
- **F10:** Wrap user SQL in an outer bound (`SELECT * FROM (<sql>) LIMIT ?`), or iterate with `stmt.iterate()` and stop at the cap. Remove the `/limit/i` regex. Add an execution budget, such as the progress handler or a worker thread, so a slow query cannot block the event loop.
- **F11 + C4:** Runtime-validate MCP tool arguments, using zod derived from the same tool specs. Read files through an open handle up to `cap + 1` bytes.

**Extra acceptance:** a 50 MB fixture read keeps peak RSS bounded; Stop and `/health` respond during a deliberately slow query.

### E7 — Cache identity, persistence & retention (F12, C5, L3)

- **F12:** Namespace the function cache by saved config id plus a non-secret config fingerprint (endpoint, model, prompt-template version). Never key on credentials.
- **C5:** Settings load tells "absent", "unreadable" and "invalid shape" apart. An unreadable file is kept as `*.corrupt-<ts>` and the problem is shown in the UI. Saves are atomic (write temp file, fsync, rename).
- **L3:** Set retention budgets (count or bytes, plus age) for the function wait queue, retained change-sets/snapshots, and the persistent cache/ledger in `db.ts`.

### E8 — Native & child-process lifecycle (L1, L2, L4)

- **L1:** In `embedded.ts`, dispose each context/session on completion, cancel, and error. Bound the model cache with LRU eviction. Validate this against the real optional native binding, not only the mock.
- **L2:** In `stdioTransport.ts`, add handshake and per-request deadlines and a pending-request cap. Start the HTTP listener without waiting for MCP sync. Handle child exit deterministically. Review how Windows shell command arguments are quoted.
- **L4:** Add IPv6 ranges to `lockdown-firewall.ps1`, verify them on a test machine, and document that child MCP binaries need their own rules.

### E9 — Docs/policy consistency (C6, C7)

- Make README and `docs/DESIGN.md` describe what is actually verified. Label cloud-named integrations as local protocol adapters, or remove them.
- Fix the stale test count in README, or drop it.
- Close this plan and the audit by noting in each which commit resolved each finding.

## Phases

| Phase | Goal | Epics | Exit criterion |
| --- | --- | --- | --- |
| **0 — Contain** | Close the demonstrated crash and bypass in days, not weeks | E0 | E0 merged; no runtime behavior change beyond the rejections it adds |
| **1 — Dependable mutations** | A reviewed change lands exactly once, on exactly the reviewed target, or fails visibly | E1, E2a, E2b | Audit probes for F1, F3, and C1 fail on `ea2f25d` and pass on the new head |
| **2 — Honest results** | A success report means the workbook is in the promised state | E3, E2c | F2, F4, and F8 probes pass; verification scope documented |
| **3 — Guarded I/O** | Every byte in or out passes one validated, bounded boundary | E4, E5, E6 | F5, F6, F7, F9, F10, F11, C2, C3, and C4 probes pass |
| **4 — Long-run hygiene** | The sidecar stays correct and bounded over long sessions | E7, E8, E9 | F12, C5, C6, and L1–L4 closed or explicitly deferred with a reason |

Phases 1 and 3 overlap mainly in `server.ts` and `shared/src/protocol.ts`, so two people can run them in parallel if they coordinate on those files. Phase 2 depends on Phase 1's state machine. E5 is the largest single epic, so start its design early.

## Tracking

| ID | Priority | Epic | Phase | Status |
| --- | --- | --- | --- | --- |
| F1 | P1 | E1 | 1 | done |
| F2 | P1 | E3 | 2 | done |
| F3 | P1 | E2a | 1 | done |
| F4 | P1 | E3 | 2 | done |
| F5 | P1 | E0 → E4 | 0 / 3 | done |
| F6 | P1 | E0 → E5 | 0 / 3 | open |
| F7 | P2 | E5 | 3 | open |
| F8 | P2 | E2c | 2 | done |
| F9 | P2 | E6 | 3 | open |
| F10 | P2 | E6 | 3 | open |
| F11 | P2 | E6 | 3 | open |
| F12 | P2 | E7 | 4 | open |
| C1 Partial apply rollback | P1 | E2b | 1 | done |
| C2 Hostname/socket pinning | P1 | E5 | 3 | open |
| C3 Anthropic `/v1` | P2 | E5 | 3 | open |
| C4 FS read allocation | P2 | E6 | 3 | open |
| C5 Settings corruption | P2 | E7 | 4 | open |
| C6 Cloud docs mismatch | P2 | E9 | 4 | open |
| C7 CI `--ignore-scripts` / README count | P3 | E0 / E9 | 0 / 4 | open |
| L1 Embedded provider cleanup | — | E8 | 4 | investigate |
| L2 MCP stdio deadlines | — | E8 | 4 | investigate |
| L3 Retention budgets | — | E7 | 4 | investigate |
| L4 Firewall IPv6 | — | E8 | 4 | investigate |
