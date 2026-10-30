# Pragmatic engineering audit — 2026-10-29

Reviewed commit: `ea2f25d` (`fixes`). This audit builds on the supplied earlier review and applies the six repository charters to concrete behavior: correct workbook targeting, predictable formulas, explicit state transitions, guarded I/O, bounded work, and representative tests.

**Assessment:** The package boundaries, shared workbook interface, staged edits, and existing verification pipeline are useful foundations. The principal weakness is that several advertised guarantees hold only on the happy path. Recovery code sometimes guesses user intent, and tests sometimes replace the production behavior that needs verification. Fix those boundaries before investing in a broader rewrite or more agent heuristics.

**Verification:** With the installed Node **22.21.1** runtime and pnpm **10.33.0**, `pnpm check`, `pnpm test`, `pnpm build`, and `pnpm smoke` all passed. Tests: **168 sidecar + 23 add-in = 191**. The initial attempt used the default Node 20.11.1 and encountered a sandbox child-process restriction; the successful run used Node 22 with the required process permissions. No dependencies were installed.

One-off probes exercised compiled sidecar code, temporary fixture files/databases, a small Office.js bridge harness, and real HTTP/WebSocket listeners on loopback. All fixture traffic stayed on this machine. The Office harness demonstrates routing decisions; it does not establish desktop Excel calculation parity. No application source was changed.

Priority meanings: **P1** should be addressed before relying on the affected safety guarantee; **P2** is a substantive correctness, availability, or maintainability issue; **P3** is housekeeping. These are engineering priorities, not CVSS ratings.

**1. P1 — An explicitly named missing worksheet silently redirects mutations.**

Evidence: [officeBridge.ts](../addin/src/officeBridge.ts), lines 368–373 and 408. `getWorksheet()` returns the active worksheet when an explicit name cannot be resolved. This resolver serves writes, clears, renames, and row/column mutations. It is used independently for snapshot reads and apply calls, so changing the active sheet can also change the actual target between phases. The emulator rejects missing sheets instead. The existing add-in test explicitly expects the fallback for `Sheet1`, but the production fallback covers every missing name.

Probe: with only an active `Payroll` worksheet, applying a write addressed to `MissingSheet!A1` sent `[[999]]` to `Payroll`. No missing-sheet error was raised.

Fix: resolve a missing sheet only during proposal construction, then show the resolved target in the preview. At execution, require the exact worksheet, allowing case-insensitive name matching. Prefer stable worksheet identity across rename/selection changes.

Acceptance: **Given** a reviewed proposal targeting `Report`, **when** `Report` is deleted before approval, **then** approval fails and `Payroll` remains untouched. Run this behavior against both executors.

**2. P1 — Formula repair changes the meaning of valid formulas.**

Evidence: [loop.ts](../sidecar/src/agent/loop.ts), lines 1277 and 1443–1455. `repairWriteRangeCells()` rewrites references based on nearby numeric blocks before schema validation. Proximity does not establish which block a total is meant to reference. The staging response also claims that all formulas are valid at line 684 without proving that claim.

Probe: stage these rows at `Sheet1!A1`:

```text
Prior A          10
Prior B          20
                 <blank>
Current A        30
Current B        40
Prior subtotal   =SUM(B1:B2)
```

The staged formula became `=SUM(B4:B5)` and evaluated to **70**, although the supplied valid formula totals **30**. The resulting formula has no calculation error, so the post-apply error check cannot detect this corruption.

Fix: preserve semantically valid formula references. Restrict automatic repair to unambiguous syntax corrections, or surface a separately explained suggestion for review. Feed ambiguous references back to the model with a structured diagnostic.

Acceptance: valid formulas referencing an earlier block, another column, absolute references, and another sheet retain their meaning. Malformed formulas produce a repair request or a visible, justified correction.

**3. P1 — Concurrent approvals can execute the same change-set twice.**

Evidence: [manager.ts](../sidecar/src/changeset/manager.ts), lines 101–129 and 159; [session.ts](../sidecar/src/session.ts), `handleApprove`; [server.ts](../sidecar/src/server.ts), lines 544–628. The status remains `staged` across asynchronous work. Each WebSocket message starts an independent async handler. The session's chat `busy` flag does not serialize mutations, and the existing double-apply test checks sequential calls only.

Probe: `Promise.all([manager.apply(id), manager.apply(id)])` for one `insert_rows` operation returned two successful results. A marker starting at A2 ended at A4; a single application should put it at A3.

Fix: acquire a per-workbook mutation lock before the first await and implement explicit applying/undoing/recovery states. Duplicate approvals should share the original result or return an already-in-progress response. Disabling the button is useful UI feedback but cannot replace server serialization.

Acceptance: simultaneous approvals insert once; approve/reject and approve/undo races produce one coherent state; different change-sets cannot interleave snapshot/apply/verify phases on the same workbook.

**4. P1 — One existing error disables regression rollback for the whole set.**

Evidence: [manager.ts](../sidecar/src/changeset/manager.ts), lines 126–159. Rollback requires `beforeErrors === 0`. Verification also treats every string starting with `#` as an error and silently continues after re-read failures.

Probe: A1 initially contained `=1/0`, and A2 contained 42. A write retaining A1's error and introducing `=2/0` at A2 returned `{ ok: true, problems: [] }`. A2 remained `#DIV/0!`.

Fix: compare before/after errors by resolved cell address and error identity. Extend the cell protocol with actual error metadata where needed. Do not report verified success when verification could not complete. Distinguish expected reads of newly created sheets from disconnections or other read failures.

Acceptance: an existing error at A1 does not exempt a new error at A2; ordinary text such as `#SKU` remains writable; failed verification yields an explicit incomplete/recovery result. Consider dependent cells outside the written ranges when defining the promised verification scope.

**5. P1 — Malformed authenticated WebSocket input escapes the server handler.**

Evidence: [server.ts](../sidecar/src/server.ts), lines 523 and 544–628. JSON parsing is followed by a TypeScript cast, not runtime validation. Accessing `msg.v` assumes a non-null object, and the detached async handler has no rejection catch. Individual sockets also lack an error handler in this block; the server uses the library's default payload limit.

Probe: an authenticated connection sending the valid JSON frame `null` raised an unhandled rejection: `Cannot read properties of null (reading 'v')`. The probe installed a temporary observer to capture it; the normal Node runtime treats this unhandled rejection as fatal. This requires a valid bearer token and is not an unauthenticated remote exploit.

Fix: define shared runtime schemas for every envelope and payload; reject malformed frames before dispatch; catch handler failures; handle socket errors; set an intentional payload limit.

Acceptance: `null`, arrays, missing payloads, invalid field types, and oversized frames close/reject only the offending connection. Health checks and a second valid client continue working.

**6. P1 — Model discovery follows redirects without validating their destinations.**

Evidence: [server.ts](../sidecar/src/server.ts), lines 714–728. The initial URL passes `assertLoopbackUrl()`, but discovery uses native `fetch()` with its default redirect handling. This is a separate gap from the earlier review's hostname-resolution concern. Chat completion requests use a different HTTP implementation.

Probe: a loopback discovery fixture redirected to an IPv4-mapped IPv6 loopback URL. `assertLoopbackUrl()` rejects that destination, yet the discovery request followed it and returned the target's model list. This demonstrates a validation bypass using exclusively loopback traffic; no external host was contacted. By source inspection, there is no validation hook for an off-host redirect either.

Fix: centralize outbound HTTP policy. Reject redirects or validate and pin every destination before connecting. Remove test-detection branches from the security boundary; use explicit transport injection or local fixtures instead. Separately implement the connected-address guarantee discussed in the original review.

Acceptance: every disallowed redirect fails before connection, across discovery, probes, chat, and custom functions; allowed local requests still work.

**7. P2 — The production HTTP client buffers SSE until EOF, defeating live streaming and its safeguards.**

Evidence: [networkGuard.ts](../sidecar/src/networkGuard.ts), lines 60–61 and 117–136; [sse.ts](../sidecar/src/providers/sse.ts), `readSseStream`. `loopbackFetch()` accumulates every response chunk and constructs a `Response` only on `end`. The SSE parser's byte cap and inactivity timer start after this buffering. There is no transport response-size cap, and socket timeouts are disabled. Tests that replace global fetch bypass this production implementation entirely.

Probe: a local server immediately wrote an SSE content event and ended 350 ms later. The fetch promise was still unavailable at 100 ms and resolved after approximately 364 ms, with the entire body already buffered.

Fix: expose a bounded readable response at headers time and propagate abort/cancel to the underlying request and response. Apply byte and inactivity limits while reading the socket. Keep a live-HTTP test proving that a delta reaches the caller before the server finishes.

Acceptance: first-token delivery precedes EOF; a stalled stream times out while the connection is open; an oversized response is rejected before it is fully retained; cancellation closes the actual socket.

**8. P2 — Undo can report success without restoring the promised state, or overwrite later work.**

Evidence: [manager.ts](../sidecar/src/changeset/manager.ts), lines 38–64 and 163–183; [protocol.ts](../shared/src/protocol.ts), staged-event payload; [ChangeSetCard.tsx](../addin/src/taskpane/components/changeset/ChangeSetCard.tsx), lines 35–47. Snapshots restore values/formulas only. `undoFidelity` is not carried in the staged event or rendered by the card. Sorts are marked `full` despite moving formatting, and every applied card offers Undo independently of later edits.

Probes:

- Add `Report`, then undo: status becomes `undone`, but `Report` still exists.
- Sort red A1=2 and blue A2=1, then undo: values return, but A1 remains blue despite `undoFidelity: 'full'`.
- Apply A1: 0→1, then 1→2; undo the first set: A1 becomes 0 while the second set remains marked `applied`.

Fix: advertise the actual restoration scope before approval and after recovery. Snapshot formatting for operations that move it, or downgrade fidelity. Enforce ordered undo or compare the current state with the applied result and surface a conflict before overwriting later work.

Acceptance: partial restoration never appears as full reversal; sort undo restores the advertised cell state; undoing an older set cannot silently erase a later change or a manual edit.

**9. P2 — The model-facing importer bypasses the approved-path boundary.**

Evidence: [loop.ts](../sidecar/src/agent/loop.ts), lines 553–628; [tabular.ts](../sidecar/src/importers/tabular.ts), lines 257–292; [server.ts](../sidecar/src/server.ts), lines 325–342. The importer accepts a model-supplied filesystem path and reads it without consulting approved paths. MCP connectors guard their reads, but this parallel tool does not. User approval of the resulting workbook write occurs after the file has already been read.

Probe: with no approved paths, a scripted model call imported a temporary `private.csv` that `PathGuard.isApproved(path, [])` rejected. Its private marker appeared in the staged cells.

Fix: distinguish file contents explicitly attached by the user from filesystem paths generated by the model. Require approved roots or a user-selected file capability before filesystem access. Apply the same rule to the HTTP parser. If unrestricted imports are intentional, document this policy exception rather than presenting approved paths as a general file-access boundary.

Acceptance: model guesses cannot read an unapproved CSV/JSON/text path; attached content and explicitly approved imports continue working.

**10. P2 — SQLite row limits are bypassed by SQL containing `LIMIT`.**

Evidence: [sqliteConnector.ts](../sidecar/src/mcp/connectors/sqliteConnector.ts), lines 223–234 and 268–275. A regex checks whether SQL contains the word `limit`; if it does, the connector skips its own limit. It then materializes the result with `stmt.all()`. A comment, string literal, nested limit, large explicit limit, or `LIMIT -1` can defeat the intended cap. Synchronous execution also blocks the event loop while the query runs.

Probe: requesting `limit: 1` with the following query returned **1,500 rows**, exceeding both the requested limit and the advertised maximum of 1,000:

```sql
WITH RECURSIVE n(x) AS (
  VALUES(1) UNION ALL SELECT x+1 FROM n WHERE x<1500
) SELECT x FROM n LIMIT 1500
```

Fix: enforce an outer result bound independent of SQL text, and stop materializing at the cap. Add a separate execution budget for expensive queries; a row cap alone cannot bound query work. Isolate long queries where they can be canceled without freezing HTTP and WebSocket service.

Acceptance: explicit large limits, comments, literals, and nested limits cannot exceed the effective row/byte cap; an expensive query does not make Stop and health checks unresponsive.

**11. P2 — Invalid MCP numeric arguments bypass the filesystem output cap too.**

Evidence: [filesystemConnector.ts](../sidecar/src/mcp/connectors/filesystemConnector.ts), lines 201–204; [manager.ts](../sidecar/src/mcp/manager.ts), `callTool`. MCP input schemas are presented to the model, but these arguments are not runtime-validated before coercion. `Number('invalid')` becomes `NaN`; the truncation comparison then evaluates false. This extends the earlier finding that file reads allocate the entire file before slicing.

Probe: `fs_read_file({ path: 'large.txt', maxBytes: 'invalid' })` on a 300 KiB fixture returned **307,200 bytes** and `truncated: false`, despite the declared **262,144-byte** cap.

Fix: runtime-validate MCP arguments with finite integer bounds. Read at most the permitted bytes plus one from an open handle; enforce output bytes separately. Share validators with the tool specifications so these cannot drift.

Acceptance: malformed numeric arguments are rejected; valid reads respect the requested and absolute caps; memory use remains bounded for a much larger fixture.

**12. P2 — Function caching conflates distinct provider configurations.**

Evidence: [batch.ts](../sidecar/src/functions/batch.ts), lines 128–131; [providers/index.ts](../sidecar/src/providers/index.ts), `createProvider`. The cache key contains only the provider's generic protocol ID, model name, and function request. The saved provider configuration ID and endpoint are not part of its identity. Two different local servers can therefore share cached results when they use the same model name.

Probe: run the same function request first against a provider returning `server A`, then against a different provider returning `server B`, both with ID `openai-compatible` and model `local-model`. The second run returned **`server A`** and called the second provider **zero times**.

Fix: include a stable, non-secret configuration identity/version in the cache namespace. Define invalidation for endpoint, model artifact, and prompt-template changes. Avoid including raw credentials in stored cache metadata.

Acceptance: switching between two local configurations with the same model name uses the selected configuration; repeated calls within an unchanged configuration still deduplicate.

**Disposition of the supplied review**

All seven earlier findings remain applicable to the reviewed commit:

| Earlier finding | Assessment and extension |
| --- | --- |
| Partial apply is not rolled back | **Confirmed by execution, P1.** Write A1 successfully, then attempt to delete the only worksheet: the call throws, A1 stays changed, and the set stays `staged`. Catch apply failures, preserve recovery state, and distinguish partial recovery from success. Source: change-set manager line 129. |
| Hostname validation is weaker than a socket guarantee | **Confirmed by inspection, P1.** `localhost` is passed to the request implementation without address pinning. No hostile DNS/hosts configuration was tested. Finding 6 adds a separate demonstrated redirect bypass. |
| Anthropic doubles `/v1` | **Confirmed by inspection, P2.** Provider factory defaults to `/v1`; the adapter appends `/v1/messages` at line 113, and discovery similarly appends `/v1/models`. Test the exact URLs, not just substring containment. |
| Filesystem reads allocate before truncating | **Confirmed by inspection, P2.** The complete file is read at filesystem connector line 202. Finding 11 additionally demonstrates bypass of the returned-byte cap. |
| Settings corruption resembles a fresh installation | **Confirmed by inspection, P2.** Config lines 47–49 silently return empty settings after read/parse errors. Saving writes the target directly at line 75. Preserve unreadable files, validate the loaded shape, distinguish absence from corruption, and save atomically. |
| Cloud documentation contradicts loopback policy | **Confirmed, P2 documentation/policy mismatch.** README and design promises exceed or contradict runtime behavior. State which integrations are local protocol adapters, and which guarantees have actually been verified. |
| CI omits `--ignore-scripts` | **Confirmed, P3 policy mismatch.** CI line 20 uses only `--frozen-lockfile`, contrary to AGENTS.md. This observation does not establish that a malicious install hook ran. The README's 162-test count is also stale. |

**Additional lifecycle work to investigate**

These are source-level follow-ups, not reproduced native-engine or long-duration failures:

- [embedded.ts](../sidecar/src/providers/embedded.ts), lines 166–200: generation creates a context/session without explicit cleanup; the model cache has no eviction bound, and backend disposal only clears references. Validate the real optional native binding's lifecycle on completion, cancellation, and errors. The green tests use a mock backend.
- [stdioTransport.ts](../sidecar/src/mcp/transports/stdioTransport.ts): JSON-RPC requests have no deadline or pending-request cap. Initial MCP synchronization is awaited before HTTP listen. A silent configured child can prevent startup; add handshake/request deadlines and deterministic child exit handling. The Windows shell also warrants a separate command/argument review.
- [batch.ts](../sidecar/src/functions/batch.ts), [manager.ts](../sidecar/src/changeset/manager.ts), and [db.ts](../sidecar/src/db.ts): the function wait queue, retained change-sets/snapshots, and persistent cache/ledger have no application retention budget. A concurrency limit or per-change snapshot cap does not bound total lifetime storage.
- [lockdown-firewall.ps1](../scripts/lockdown-firewall.ps1): the configured blocked ranges cover IPv4 only. Verify non-loopback IPv6 separately before relying on the script as a complete egress restriction. Child MCP binaries need their own restriction boundary. No firewall state was changed or tested in this audit.

**Recommended implementation order**

1. Make workbook mutations dependable: exact target resolution, per-workbook serialization, explicit partial-apply recovery, and conflict-aware undo.
2. Preserve formula meaning and replace the global error-count shortcut with address-level verification. Document exactly what dependent-cell checks cover.
3. Consolidate I/O boundaries: validated protocol envelopes, one bounded HTTP transport, redirect/address enforcement, approved file capabilities, and runtime MCP validation.
4. Add representative regression coverage: parallel approvals, preserved valid formulas, workbook changes between preview/apply/undo, malformed frames, live local SSE, redirect rejection, and bounded reads/queries. Retain the existing fast unit tests and smoke test.
5. Tighten cache identity, native/process cleanup, retention budgets, atomic settings persistence, and documentation/CI consistency.

The useful pragmatic principle here is to make guarantees executable at their owning boundary. Types and prompts describe the intended contract; real execution must enforce it. The existing modular structure can support these repairs without a wholesale architecture replacement.
