import { existsSync } from 'node:fs';
import { basename, extname, resolve } from 'node:path';
import { zodToJsonSchema } from 'zod-to-json-schema';
import {
  toolSchemas,
  isWriteTool,
  colToIndex,
  indexToCol,
  formatA1,
  parseA1,
  type ChangeOp,
  type ToolName,
  type ChangePreviewItem,
  type UndoFidelity,
  type WorkbookMap,
} from '@excelai/shared';
import type { ChatMessage, Provider, ToolSpec } from '../providers/types.js';
import { trimRuminatingReasoning, withRetry } from '../providers/index.js';
import { extractThinking, extractToolCalls } from '../providers/embedded.js';
import type { WorkbookExecutor } from '../workbook/executor.js';
import type { ChangeSetManager } from '../changeset/manager.js';
import { resolveSheetTargets } from '../changeset/targets.js';
import { buildSystemPrompt, emulatedToolInstructions } from './prompts.js';
import {
  generateImportChangeOps,
  loadAndParseLocalFile,
  parseMarkdownTable,
  parseTabularContent,
} from '../importers/tabular.js';
import type { McpManager } from '../mcp/manager.js';
import { PathGuard } from '../mcp/pathGuard.js';

export type AgentEvent =
  | { kind: 'thinking_delta'; text: string }
  | { kind: 'text_delta'; text: string }
  | { kind: 'thinking'; text: string; durationMs?: number }
  | { kind: 'text'; text: string }
  | { kind: 'tool_use'; tool: string; summary: string }
  | {
      kind: 'changeset_staged';
      changeSetId: string;
      preview: ChangePreviewItem[];
      undoFidelity: UndoFidelity;
    }
  | { kind: 'error'; message: string }
  | { kind: 'stopped'; reason?: string };

export interface AgentTurnOptions {
  provider: Provider;
  executor: WorkbookExecutor;
  changeSets: ChangeSetManager;
  workbookName: string;
  history: ChatMessage[];
  userMessage: string;
  onEvent: (ev: AgentEvent) => void;
  onUsage?: (usage: { inputTokens: number; outputTokens: number }) => void;
  maxIterations?: number;
  abortSignal?: AbortSignal;
  mcpManager?: McpManager;
  /** Folders the user approved; import_local_data may only open files inside them. */
  approvedPaths?: readonly string[];
}

export interface AgentTurnResult {
  history: ChatMessage[];
  stagedChangeSetId?: string;
}

const MAX_TOOL_RESULT_CHARS = 20_000;
const MAX_REPAIR_ATTEMPTS_PER_CALL = 1; // a malformed call gets one structured error back per attempt

const TOOL_ALIASES: Record<string, ToolName> = {
  get_active_sheet: 'get_workbook_map',
  get_sheets: 'get_workbook_map',
  list_sheets: 'get_workbook_map',
  get_sheet_names: 'get_workbook_map',
  get_workbook_info: 'get_workbook_map',
};

function normalizeToolName(name: string): string {
  const trimmed = name.trim();
  return TOOL_ALIASES[trimmed] ?? trimmed;
}

export function buildToolSpecs(): ToolSpec[] {
  return (Object.keys(toolSchemas) as ToolName[]).map((name) => ({
    name,
    description: toolSchemas[name].description,
    inputSchema: zodToJsonSchema(toolSchemas[name].parameters, { $refStrategy: 'none' }) as Record<
      string,
      unknown
    >,
  }));
}

/**
 * One user-visible agent turn: a bounded tool-calling loop in which every
 * tool argument is schema-validated before it does anything, write tools are
 * staged into a change-set (never applied), and malformed calls get a
 * structured error fed back so the model can repair itself.
 */
export async function runAgentTurn(options: AgentTurnOptions): Promise<AgentTurnResult> {
  const {
    provider,
    executor,
    changeSets,
    workbookName,
    onEvent,
    onUsage,
    maxIterations = 24,
    abortSignal,
  } = options;

  let cachedMap: WorkbookMap | null = null;
  const getCachedMap = async (): Promise<WorkbookMap> => {
    if (!cachedMap) {
      cachedMap = await executor.getWorkbookMap();
    }
    return cachedMap;
  };

  let prefetchTimer: ReturnType<typeof setTimeout> | undefined;
  try {
    cachedMap = await Promise.race([
      executor.getWorkbookMap(),
      new Promise<null>((resolvePrefetch) => {
        prefetchTimer = setTimeout(() => resolvePrefetch(null), 3_000);
      }),
    ]);
  } catch {
    // Non-fatal: if pre-fetching the map fails, tools can still fetch it on demand
  } finally {
    if (prefetchTimer !== undefined) clearTimeout(prefetchTimer);
  }

  let system = buildSystemPrompt(workbookName, cachedMap);
  const mcpSpecs = options.mcpManager?.getToolSpecs() ?? [];
  const specs = [...buildToolSpecs(), ...mcpSpecs];
  if (!provider.supportsTools) {
    const toolList = specs
      .map((s) => `- ${s.name}: ${s.description}\n  args schema: ${JSON.stringify(s.inputSchema)}`)
      .join('\n');
    system += `\n\n${emulatedToolInstructions(toolList)}`;
  }

  const messages: ChatMessage[] = [
    { role: 'system', content: system },
    ...options.history,
    { role: 'user', content: options.userMessage },
  ];

  const pendingOps: ChangeOp[] = [];
  const seenCalls = new Map<string, { count: number; lastResult: string }>();
  const writeRangeStartHistory = new Set<string>();
  // Formulas flagged this turn (cell -> formula); sent again unchanged, they count as confirmed.
  const flaggedFormulas = new Map<string, string>();
  const duplicateAssistantIndices = new Set<number>();
  const duplicateToolCallIds = new Set<string>();
  const recoveryMessageIndices = new Set<number>();
  let forceTextReplyNext = false;
  let emittedAnything = false;
  let reasoningRecoveryCount = 0;
  let deferredWriteRetryCount = 0;
  const MAX_REASONING_RECOVERY_ATTEMPTS = 2;
  const MAX_DEFERRED_WRITE_RETRIES = 2;
  const RECOVERY_NUDGE =
    '[System note: Your previous completion ran out of tokens or hit a JSON formatting loop before finishing the tool call. Do not deliberate, recount rows, or pretty-print — immediately emit your next tool call (e.g. write_range starting at "Sheet1!A1") as compact single-line JSON with no extra whitespace, or reply with your final summary concisely.]';
  const emit = (ev: AgentEvent): void => {
    if (ev.kind !== 'thinking' && ev.kind !== 'thinking_delta') {
      emittedAnything = true;
    }
    onEvent(ev);
  };

  for (let iteration = 0; iteration < maxIterations; iteration++) {
    if (abortSignal?.aborted) {
      emit({ kind: 'stopped', reason: 'Generation stopped by user.' });
      return { history: options.history };
    }

    const disableToolsForStep =
      forceTextReplyNext || (iteration === maxIterations - 1 && iteration > 0);
    forceTextReplyNext = false;

    const stepMaxTokens = reasoningRecoveryCount > 0 ? 16384 : 8192;
    const stepStartedAt = Date.now();
    let streamedThinkingInStep = false;
    let announcedToolStream = false;

    let response;
    const providerMessages = prepareMessagesForProvider(messages, 1 + options.history.length);
    try {
      response = await withRetry(() =>
        provider.chat({
          messages: providerMessages,
          tools: provider.supportsTools && !disableToolsForStep ? specs : undefined,
          temperature: reasoningRecoveryCount > 0 || deferredWriteRetryCount > 0 ? 0.2 : 0,
          reasoningEffort: 'low',
           // Generous budget for reasoning + large write_range tool calls,
          // expanded further when recovering from reasoning token exhaustion.
          maxTokens: stepMaxTokens,
          abortSignal,
          onDelta: (delta) => {
            if (abortSignal?.aborted) return;
            if (delta.reasoning) {
              streamedThinkingInStep = true;
              emit({ kind: 'thinking_delta', text: delta.reasoning });
            }
            if (delta.toolCallName && !announcedToolStream && !streamedThinkingInStep) {
              announcedToolStream = true;
              emit({
                kind: 'thinking_delta',
                text: `Preparing ${delta.toolCallName} tool call…`,
              });
            }
            if (delta.text) {
              emit({ kind: 'text_delta', text: delta.text });
            }
          },
        }),
      );
    } catch (e) {
      if (
        abortSignal?.aborted ||
        (e instanceof Error && (e.name === 'AbortError' || e.message.toLowerCase().includes('abort')))
      ) {
        emit({ kind: 'stopped', reason: 'Generation stopped by user.' });
        return { history: options.history };
      }
      throw e;
    }

    const stepDurationMs = Math.max(1, Date.now() - stepStartedAt);

    if (abortSignal?.aborted) {
      emit({ kind: 'stopped', reason: 'Generation stopped by user.' });
      return { history: options.history };
    }

    onUsage?.(response.usage);

    let toolCalls = response.toolCalls;
    let assistantText = response.text;
    let inlineReasoning = '';
    if (/<(?:think|thinking)>/i.test(assistantText)) {
      const extractedThinking = extractThinking(assistantText);
      assistantText = extractedThinking.cleanText;
      inlineReasoning = extractedThinking.reasoning;
    }
    const thinkingText = [response.reasoning?.trim(), inlineReasoning.trim()]
      .filter(Boolean)
      .join('\n\n');
    if (thinkingText.length > 0) {
      emit({ kind: 'thinking', text: thinkingText, durationMs: stepDurationMs });
    } else if (announcedToolStream) {
      emit({
        kind: 'thinking',
        text: `Prepared ${toolCalls[0]?.name ?? 'tool'} call.`,
        durationMs: stepDurationMs,
      });
    }

    // Fallback tool-call extraction: local models on OpenAI-compatible servers
    // (LM Studio, Ollama, llama.cpp) sometimes emit <tool_call>...</tool_call>
    // or JSON tool blocks in `content` with `tool_calls: []`.
    if (toolCalls.length === 0 && assistantText.trim()) {
      const hadToolCallTag = /<(?:tool_call|function=|arg_key>)/i.test(assistantText);
      if (hadToolCallTag || !provider.supportsTools) {
        const extracted = extractToolCalls(assistantText);
        if (extracted.toolCalls.length > 0) {
          toolCalls = extracted.toolCalls;
          assistantText = extracted.cleanText;
        } else if (hadToolCallTag) {
          assistantText = extracted.cleanText;
          if (!assistantText.trim()) {
            toolCalls = [
              {
                id: `malformed_${iteration}`,
                name: '__malformed_tool_call__',
                argsJson: '{}',
              },
            ];
          }
        }
      }
      if (toolCalls.length === 0) {
        const parsed = tryParseEmulatedCall(assistantText);
        if (parsed) {
          toolCalls = [
            { id: `emu_${iteration}`, name: parsed.tool, argsJson: JSON.stringify(parsed.args) },
          ];
          assistantText = '';
        }
      }
    }

    if (toolCalls.length === 0) {
      // Rescue complete tool-call JSON payloads drafted inside reasoning_content when the
      // model hit max_tokens before emitting the formal tool call.
      if (!assistantText.trim() && response.stopReason === 'max_tokens' && thinkingText.trim()) {
        const rescuedFromThinking = tryExtractToolCallFromReasoning(thinkingText, iteration);
        if (rescuedFromThinking) {
          toolCalls = [rescuedFromThinking];
        }
      }
    }

    // Guard against "all-talk / no-tool" replies where the model refuses to overwrite a prior turn's
    // staged change-set, promises to stage the reorganized layout "next", or claims it staged changes
    // in plain text when pendingOps is still empty.
    if (
      toolCalls.length === 0 &&
      pendingOps.length === 0 &&
      assistantText.trim() &&
      !disableToolsForStep &&
      deferredWriteRetryCount < MAX_DEFERRED_WRITE_RETRIES &&
      iteration + 1 < maxIterations &&
      shouldRetryDeferredOrUnstagedWrite(assistantText, options.userMessage, options.history)
    ) {
      deferredWriteRetryCount++;
      emit({
        kind: 'thinking_delta',
        text: '\nStaging requested layout via write_range…\n',
      });
      recoveryMessageIndices.add(messages.length);
      messages.push({ role: 'assistant', content: assistantText });
      const nudge = buildDeferredWriteRetryNudge(
        options.userMessage,
        assistantText,
        options.history,
      );
      recoveryMessageIndices.add(messages.length);
      messages.push({ role: 'user', content: nudge });
      continue;
    }

    if (assistantText.trim()) emit({ kind: 'text', text: assistantText });

    if (toolCalls.length === 0) {
      if (!assistantText.trim()) {
        if (
          response.stopReason === 'max_tokens' &&
          reasoningRecoveryCount < MAX_REASONING_RECOVERY_ATTEMPTS &&
          iteration + 1 < maxIterations
        ) {
          reasoningRecoveryCount++;
          const nudgeContent = buildReasoningRecoveryNudge(
            thinkingText,
            options.userMessage,
            RECOVERY_NUDGE,
          );
          const lastIdx = messages.length - 1;
          if (lastIdx >= 0 && recoveryMessageIndices.has(lastIdx)) {
            messages[lastIdx] = { role: 'user', content: nudgeContent };
          } else {
            recoveryMessageIndices.add(messages.length);
            messages.push({ role: 'user', content: nudgeContent });
          }
          continue;
        }
        // A turn must never end in silence: explain empty replies.
        emit({
          kind: 'error',
          message:
            response.stopReason === 'max_tokens'
              ? 'The model ran out of output tokens before producing a reply (it may have spent them on internal reasoning). Try a shorter request, or a model with a larger output budget.'
              : 'The model returned an empty reply. Try rephrasing, or test the model in Settings.',
        });
      } else {
        messages.push({ role: 'assistant', content: assistantText, toolCalls });
      }
      break;
    }

    reasoningRecoveryCount = 0;
    const assistantMsgIdx = messages.length;
    messages.push({ role: 'assistant', content: assistantText, toolCalls });

    const priorWriteStarts = new Set(writeRangeStartHistory);
    let allCallsInStepWereDuplicates = true;
    for (const call of toolCalls) {
      if (abortSignal?.aborted) {
        emit({ kind: 'stopped', reason: 'Generation stopped by user.' });
        return { history: options.history };
      }
      const canonicalName = normalizeToolName(call.name);
      const callSig = `${canonicalName}:${call.argsJson.trim()}`;
      const prevCall = seenCalls.get(callSig);
      // Re-sending a write whose formulas were flagged confirms them, so it must run again.
      const confirmsFlaggedWrite =
        prevCall !== undefined && (safeParseJson(prevCall.lastResult) as { diagnostics?: unknown } | null)?.diagnostics !== undefined;
      if (prevCall && !confirmsFlaggedWrite) {
        prevCall.count++;
        forceTextReplyNext = true;
        duplicateToolCallIds.add(call.id);
        const dupResult = JSON.stringify({
          ok: true,
          duplicateCall: true,
          note: `STOP: You already called ${canonicalName} with these exact arguments in this turn. Do NOT call ${canonicalName} with the same arguments again. Use previousResult below to proceed — call write_range / format_range now if creating or updating cells, or reply in plain text to answer the user.`,
          previousResult: safeParseJson(prevCall.lastResult),
        });
        messages.push({ role: 'tool', toolCallId: call.id, content: dupResult });
        continue;
      }

      allCallsInStepWereDuplicates = false;
      let result: string;
      try {
        result = await executeToolCall(
          canonicalName,
          call.argsJson,
          executor,
          workbookName,
          getCachedMap,
          messages,
          pendingOps,
          emit,
          flaggedFormulas,
          options.approvedPaths ?? [],
          options.mcpManager,
          abortSignal,
          (normalizedArgsJson, writeStartKey) => {
            call.argsJson = normalizedArgsJson;
            if (writeStartKey) {
              if (priorWriteStarts.has(writeStartKey)) {
                // Model already wrote to this start cell in an earlier step and is re-staging;
                // accept this revision and force the final text reply next so it cannot loop indefinitely.
                forceTextReplyNext = true;
              }
              writeRangeStartHistory.add(writeStartKey);
            }
          },
        );
      } catch (e) {
        if (
          abortSignal?.aborted ||
          (e instanceof Error && (e.name === 'AbortError' || e.message.toLowerCase().includes('abort')))
        ) {
          emit({ kind: 'stopped', reason: 'Generation stopped by user.' });
          return { history: options.history };
        }
        throw e;
      }
      if (abortSignal?.aborted) {
        emit({ kind: 'stopped', reason: 'Generation stopped by user.' });
        return { history: options.history };
      }
      seenCalls.set(callSig, { count: 1, lastResult: result });
      messages.push({ role: 'tool', toolCallId: call.id, content: result });
    }

    if (allCallsInStepWereDuplicates && !assistantText.trim()) {
      duplicateAssistantIndices.add(assistantMsgIdx);
    }
  }

  // Loop exhausted while the model was still calling tools.
  if (messages[messages.length - 1]?.role === 'tool') {
    emit({
      kind: 'error',
      message: `Stopped after ${maxIterations} steps without a final answer. You can ask the model to continue.`,
    });
  }

  const result: AgentTurnResult = {
    // Persist everything after the system prompt as conversation history,
    // excluding temporary recovery nudges and duplicate tool-call loops.
    history: messages.filter(
      (m, idx) =>
        idx > 0 &&
        !recoveryMessageIndices.has(idx) &&
        !duplicateAssistantIndices.has(idx) &&
        !(m.role === 'tool' && duplicateToolCallIds.has(m.toolCallId)),
    ),
  };

  if (pendingOps.length > 0) {
    try {
      const cs = await changeSets.stage(pendingOps);
      const preview = await changeSets.preview(cs.id);
      emit({ kind: 'changeset_staged', changeSetId: cs.id, preview, undoFidelity: cs.undoFidelity });
      result.stagedChangeSetId = cs.id;
    } catch (e) {
      emit({
        kind: 'error',
        message: `The proposed changes could not be staged: ${e instanceof Error ? e.message : String(e)}`,
      });
    }
  }

  if (!emittedAnything) {
    // Belt and braces: whatever happened above, the user gets a signal.
    onEvent({
      kind: 'error',
      message: 'The model produced no visible output for this turn. Check the sidecar console for details.',
    });
  }

  return result;
}

async function executeToolCall(
  name: string,
  argsJson: string,
  executor: WorkbookExecutor,
  workbookName: string,
  getCachedMap: () => Promise<WorkbookMap>,
  messages: ChatMessage[],
  pendingOps: ChangeOp[],
  onEvent: (ev: AgentEvent) => void,
  flaggedFormulas: Map<string, string>,
  approvedPaths: readonly string[],
  mcpManager?: McpManager,
  abortSignal?: AbortSignal,
  onArgsNormalized?: (normalizedArgsJson: string, writeStartKey?: string) => void,
): Promise<string> {
  if (mcpManager?.hasTool(name)) {
    let rawArgs: Record<string, unknown> = {};
    try {
      rawArgs = argsJson.trim() === '' ? {} : JSON.parse(argsJson);
    } catch {
      return errorResult('Tool arguments were not valid JSON. Re-send the call with valid JSON arguments.');
    }
    const summary = JSON.stringify(rawArgs).slice(0, 100);
    onEvent({ kind: 'tool_use', tool: name, summary });

    const mcpRes = await mcpManager.callTool(name, rawArgs, { abortSignal });
    if (mcpRes.changeOps && mcpRes.changeOps.length > 0) {
      pendingOps.push(...mcpRes.changeOps);
    }
    return mcpRes.resultText;
  }

  if (!(name in toolSchemas)) {
    const available = [
      ...Object.keys(toolSchemas),
      ...(mcpManager ? mcpManager.getToolSpecs().map((s) => s.name) : []),
    ];
    return errorResult(`Unknown tool "${name}". Available tools: ${available.join(', ')}`);
  }
  const toolName = name as ToolName;

  let rawArgs: unknown;
  try {
    rawArgs = argsJson.trim() === '' ? {} : JSON.parse(argsJson);
  } catch {
    return errorResult('Tool arguments were not valid JSON. Re-send the call with valid JSON arguments.');
  }

  let formulaCorrections: FormulaCorrection[] = [];
  let confirmedFormulas: FormulaDiagnostic[] = [];
  if (typeof rawArgs === 'object' && rawArgs !== null) {
    const rec = rawArgs as Record<string, unknown>;
    if (typeof rec.range === 'string' && rec.range.trim().length > 0) {
      rec.range = await normalizeRangeForWorkbook(
        rec.range,
        workbookName,
        getCachedMap,
        pendingOps,
      );
    }
    if (toolName === 'write_range' && Array.isArray(rec.cells)) {
      const repaired = repairWriteRangeCells(
        typeof rec.range === 'string' ? rec.range : 'Sheet1!A1',
        rec.cells as unknown[],
      );
      rec.cells = repaired.cells;
      formulaCorrections = repaired.corrections;
      // Never rewrite a reference: send it back, unless the model already saw it flagged and kept it.
      confirmedFormulas = repaired.diagnostics.filter((d) => flaggedFormulas.get(d.cell) === d.formula);
      const unresolved = repaired.diagnostics.filter((d) => !confirmedFormulas.includes(d));
      if (unresolved.length > 0) {
        for (const d of unresolved) flaggedFormulas.set(d.cell, d.formula);
        return JSON.stringify({
          ok: false,
          staged: false,
          error:
            'Nothing was staged: these formulas refer to blank or label cells of the block being written. ' +
            'Re-send write_range with each formula corrected (candidates are suggestions, not certainties). ' +
            'If a formula is right as written, re-send it unchanged and it will be staged as is.',
          diagnostics: unresolved,
          ...(formulaCorrections.length > 0 ? { corrections: formulaCorrections } : {}),
        });
      }
    }
  }

  const parsed = toolSchemas[toolName].parameters.safeParse(rawArgs);
  if (!parsed.success) {
    const issues = parsed.error.issues
      .map((i) => `${i.path.join('.') || '(root)'}: ${i.message}`)
      .join('; ');
    return errorResult(`Invalid arguments for ${toolName}: ${issues}`);
  }
  const args = parsed.data as Record<string, unknown>;

  onEvent({ kind: 'tool_use', tool: toolName, summary: summarize(toolName, args) });

  try {
    if (isWriteTool(toolName)) {
      if (toolName === 'import_local_data') {
        const filePath = args.filePath as string;
        const startCell = (args.startCell as string | undefined) ?? 'A1';
        const hasHeaders = (args.hasHeaders as boolean | undefined) ?? true;
        const reason = (args.reason as string | undefined) ?? `Import data from ${filePath}`;
        let targetSheet = args.targetSheet as string | undefined;
        if (!targetSheet) {
          const map = await getCachedMap();
          targetSheet = map.sheets[0]?.name ?? 'Sheet1';
        }

        // Attached content needs no filesystem access; a path on disk must sit in an approved folder.
        const approved = PathGuard.isApproved(filePath, [...approvedPaths]);
        if (!approved || !existsSync(resolve(filePath))) {
          const attached = findAttachedDocument(messages, filePath);
          if (attached) {
            const ext = extname(attached.name).toLowerCase();
            const isNarrativeDoc = ext === '.md' || ext === '.markdown' || ext === '.txt';
            if (isNarrativeDoc) {
              const mdRows = parseMarkdownTable(attached.content);
              const nonTableLines = attached.content
                .split(/\r?\n/)
                .map((l) => l.trim())
                .filter((l) => l.length > 0 && !l.includes('|') && !l.startsWith('#'));

              if (mdRows.length < 2 || nonTableLines.length > mdRows.length) {
                return JSON.stringify({
                  ok: true,
                  alreadyInContext: true,
                  documentName: attached.name,
                  targetSheet,
                  startCell,
                  note: `The full contents of "${attached.name}" are already provided inline in the user's message under [Attached Document: ${attached.name}]. It is a specification/narrative document rather than a raw tabular file on disk. Do NOT call import_local_data or search for "${attached.name}" again — read the requirements directly from the user message and call write_range and format_range on "${targetSheet}" starting at ${startCell} to build the spreadsheet.`,
                });
              }
            }

            const format =
              ext === '.tsv'
                ? 'tsv'
                : ext === '.json'
                  ? 'json'
                  : ext === '.md' || ext === '.markdown'
                    ? 'markdown'
                    : ext === '.csv'
                      ? 'csv'
                      : 'auto';
            const rows = parseTabularContent(attached.content, format);
            const totalRows = rows.length;
            const totalCols = rows[0]?.length ?? 0;
            const ops = generateImportChangeOps({
              rows,
              targetSheet,
              startCell,
              hasHeaders,
              reason,
            });
            pendingOps.push(...ops);
            return JSON.stringify({
              ok: true,
              staged: true,
              format,
              rowsCount: totalRows,
              colsCount: totalCols,
              targetRange: ops[0]?.kind === 'write_range' ? ops[0].range : undefined,
              note: `Imported ${totalRows} rows × ${totalCols} columns from attached "${attached.name}" into pending change-set. Not applied yet until user approves.`,
            });
          }

          if (!approved) {
            return errorResult(
              `"${filePath}" is not in a folder you have approved for import, so it was not read. If the user attached the document in chat, its contents are already in the user message under [Attached Document: ...] — use write_range directly. Otherwise ask the user to attach the file or approve its folder in Settings; do not retry with guessed paths.`,
            );
          }
          return errorResult(
            `File not found on local disk: "${filePath}". Do not guess directory paths (such as /mnt/data/) or retry import_local_data with guessed paths. If the user attached a document in chat, its full contents are already in the user message under [Attached Document: ...] — use write_range directly.`,
          );
        }

        const { rows, totalRows, totalCols, format } = loadAndParseLocalFile(filePath, approvedPaths);
        const ops = generateImportChangeOps({
          rows,
          targetSheet,
          startCell,
          hasHeaders,
          reason,
        });
        pendingOps.push(...ops);
        return JSON.stringify({
          ok: true,
          staged: true,
          format,
          rowsCount: totalRows,
          colsCount: totalCols,
          targetRange: ops[0]?.kind === 'write_range' ? ops[0].range : undefined,
          note: `Imported ${totalRows} rows × ${totalCols} columns staged into pending change-set. Not applied yet until user approves.`,
        });
      }
      const op = toChangeOp(toolName, args);
      const sheetProblem = await findSheetProblem([...pendingOps, op], getCachedMap);
      if (sheetProblem) return errorResult(sheetProblem);
      if (op.kind === 'write_range') {
        const opAddr = parseA1(op.range);
        const startKey = `${opAddr.sheet.toLowerCase()}!R${opAddr.startRow}C${opAddr.startCol}`;
        onArgsNormalized?.(JSON.stringify(args), startKey);
        const startCellA1 = formatA1({
          sheet: opAddr.sheet,
          startRow: opAddr.startRow,
          startCol: opAddr.startCol,
          endRow: opAddr.startRow,
          endCol: opAddr.startCol,
        });
        const existingIdx = pendingOps.findIndex((p) => {
          if (p.kind !== 'write_range') return false;
          const pAddr = parseA1(p.range);
          return (
            pAddr.sheet.toLowerCase() === opAddr.sheet.toLowerCase() &&
            pAddr.startRow === opAddr.startRow &&
            pAddr.startCol === opAddr.startCol
          );
        });
        const replacedPriorDraft = existingIdx >= 0;
        if (replacedPriorDraft) {
          pendingOps[existingIdx] = op;
        } else {
          pendingOps.push(op);
        }
        const rowsStaged = op.cells.length;
        const colsStaged = op.cells[0]?.length ?? 0;
        return JSON.stringify({
          ok: true,
          staged: true,
          targetRange: op.range,
          rowsStaged,
          colsStaged,
          replacedPriorDraft,
          stagedOpsCount: pendingOps.length,
          ...(formulaCorrections.length > 0 ? { corrections: formulaCorrections } : {}),
          ...(confirmedFormulas.length > 0 ? { stagedAsWritten: confirmedFormulas.map((d) => d.cell) } : {}),
          note: `Staged ${rowsStaged}×${colsStaged} table at ${op.range} into the pending change-set for this turn${replacedPriorDraft ? ` (replaced prior draft at ${startCellA1})` : ''}.${formulaCorrections.length > 0 ? ' Formula syntax was corrected as listed in corrections; references were not changed.' : ''} Formulas are checked only for syntax here; the user reviews their results before applying. For this turn, you may now call format_range once if needed, or reply to the user with a concise summary.`,
        });
      }
      pendingOps.push(op);
      return JSON.stringify({
        ok: true,
        staged: true,
        stagedOpsCount: pendingOps.length,
        note: 'Staged into the pending change-set. Not applied yet; reads will not reflect it until the user approves.',
      });
    }
    switch (toolName) {
      case 'get_workbook_map': {
        const map = await getCachedMap();
        return JSON.stringify({ ok: true, result: map });
      }
      case 'read_range': {
        const cells = await executor.readRange(args.range as string);
        const json = JSON.stringify({ ok: true, result: cells });
        if (json.length > MAX_TOOL_RESULT_CHARS) {
          return errorResult(
            `Result too large (${json.length} chars). Read a smaller, more targeted range.`,
          );
        }
        return json;
      }
      case 'find': {
        const matches = await executor.find(args.query as string, args.sheet as string | undefined);
        return JSON.stringify({ ok: true, result: matches });
      }
      default:
        return errorResult(`Tool ${toolName} is not executable here.`);
    }
  } catch (e) {
    return errorResult(e instanceof Error ? e.message : String(e));
  }
}

/**
 * Searches conversation user messages for `[Attached Document: <name>...]\n```\n<content>\n````
 * matching the requested filePath (by basename).
 */
function findAttachedDocument(
  messages: ChatMessage[],
  requestedPath: string,
): { name: string; content: string } | null {
  const reqBase = basename(requestedPath.replace(/\\/g, '/')).trim().toLowerCase();
  const docRegex = /\[Attached Document:\s*([^\n\]]+?)\]\r?\n```[^\r\n]*\r?\n([\s\S]*?)\r?\n```/g;
  const allFound: { name: string; content: string }[] = [];

  for (let i = messages.length - 1; i >= 0; i--) {
    const msg = messages[i];
    if (!msg || msg.role !== 'user') continue;
    docRegex.lastIndex = 0;
    let m: RegExpExecArray | null;
    while ((m = docRegex.exec(msg.content)) !== null) {
      const rawHeader = (m[1] ?? '').trim();
      const name = rawHeader.replace(/\s*\(full contents.*$/i, '').trim();
      const content = m[2] ?? '';
      if (!name) continue;
      const docBase = basename(name.replace(/\\/g, '/')).toLowerCase();
      if (docBase === reqBase || name.toLowerCase() === requestedPath.trim().toLowerCase()) {
        return { name, content };
      }
      allFound.push({ name, content });
    }
  }

  if (allFound.length === 1) {
    return allFound[0]!;
  }
  return null;
}

/**
 * Prepares conversation messages for the LLM provider without mutating `messages`:
 * 1. Deduplicates repeated `[Attached Document: <name>...]` blocks across multi-turn
 *    history so re-attaching the same file on Turn 2+ does not double/triple prompt tokens
 *    and stall local model context windows.
 * 2. Compacts oversized `write_range` `argsJson` payloads from prior turns (`idx < priorHistoryEndIdx`)
 *    so subsequent turns retain the target range, top/bottom rows, and reason without re-sending
 *    thousands of tokens of historical 2D cell arrays.
 */
function prepareMessagesForProvider(
  messages: ChatMessage[],
  priorHistoryEndIdx: number,
): ChatMessage[] {
  const seenDocsFromNewest = new Set<string>();
  let keptLatestPriorWriteRange = false;
  const docBlockRe = /\[Attached Document:\s*([^\n\]]+?)\]\r?\n```[^\r\n]*\r?\n[\s\S]*?\r?\n```/g;

  const prepared: ChatMessage[] = new Array(messages.length);
  for (let i = messages.length - 1; i >= 0; i--) {
    const msg = messages[i]!;
    if (msg.role === 'user' && msg.content.includes('[Attached Document:')) {
      const updatedContent = msg.content.replace(docBlockRe, (fullBlock, rawHeader: string) => {
        const cleanName = rawHeader.replace(/\s*\(full contents.*$/i, '').trim();
        const key = cleanName.toLowerCase();
        if (!key) return fullBlock;
        if (seenDocsFromNewest.has(key)) {
          return `[Attached Document: ${cleanName} — full contents included in latest message below]`;
        }
        seenDocsFromNewest.add(key);
        return fullBlock;
      });
      prepared[i] = updatedContent !== msg.content ? { ...msg, content: updatedContent } : msg;
      continue;
    }

    if (msg.role === 'tool' && i < priorHistoryEndIdx) {
      if (msg.content.includes('"staged":true') || msg.content.includes('Do NOT call write_range')) {
        try {
          const parsed = JSON.parse(msg.content) as Record<string, unknown>;
          if (typeof parsed.note === 'string') {
            parsed.note =
              'Completed in a previous turn. If the user asks to reorganize, revise, or replace this table in the current turn (or confirms with "yes" / "sounds good" / "let\'s do it"), you MUST call write_range again in this turn (e.g. starting at Sheet1!A1) to stage the new version.';
            prepared[i] = { ...msg, content: JSON.stringify(parsed) };
            continue;
          }
        } catch {
          // ignore
        }
      }
      prepared[i] = msg;
      continue;
    }

    if (msg.role === 'assistant' && i < priorHistoryEndIdx) {
      let cleanedContent = msg.content;
      if (cleanedContent) {
        cleanedContent = cleanedContent
          .replace(
            /[^.!\n]*(?:kept the staged version|keeping the prior staged version|previous version is already staged|won['’]t overwrite it again)[^.!\n]*[.!]?\s*/gi,
            '',
          )
          .trim();
      }

      if (msg.toolCalls && msg.toolCalls.length > 0) {
        let changed = cleanedContent !== msg.content;
        const compactedCalls = msg.toolCalls.map((tc) => {
          if (tc.name !== 'write_range') return tc;
          // Always preserve the most recent prior write_range call in full (up to 12,000 chars)
          // so follow-up turns ("reorganize it left to right") have all accounts and values from Turn 1.
          if (!keptLatestPriorWriteRange && tc.argsJson.length <= 12000) {
            keptLatestPriorWriteRange = true;
            return tc;
          }
          keptLatestPriorWriteRange = true;
          if (tc.argsJson.length <= 1200) return tc;
          try {
            const parsed = JSON.parse(tc.argsJson) as Record<string, unknown>;
            if (!Array.isArray(parsed.cells) || parsed.cells.length <= 8) return tc;
            const totalRows = parsed.cells.length;
            const omitted = totalRows - 6;
            const compactCells = [
              ...parsed.cells.slice(0, 3),
              [`... (${omitted} middle rows omitted from older turn history; total ${totalRows} rows)`],
              ...parsed.cells.slice(-3),
            ];
            changed = true;
            return {
              ...tc,
              argsJson: JSON.stringify({
                ...parsed,
                cells: compactCells,
              }),
            };
          } catch {
            return tc;
          }
        });
        prepared[i] = changed
          ? { ...msg, content: cleanedContent, toolCalls: compactedCalls }
          : msg;
        continue;
      }

      prepared[i] =
        cleanedContent !== msg.content ? { ...msg, content: cleanedContent } : msg;
      continue;
    }

    prepared[i] = msg;
  }

  return prepared;
}

function toChangeOp(name: ToolName, args: Record<string, unknown>): ChangeOp {
  switch (name) {
    case 'write_range': {
      const rawCells = args.cells as (string | number | boolean | null | { formula: string })[][];
      const addr = parseA1(args.range as string);
      const maxCols = Math.max(
        addr.endCol - addr.startCol + 1,
        ...rawCells.map((r) => (r.length > 0 ? r.length : 1)),
      );
      const cells = rawCells.map((row) => {
        const mapped = row.map((cell) =>
          typeof cell === 'object' && cell !== null && 'formula' in cell
            ? { formula: cell.formula }
            : { value: cell },
        );
        while (mapped.length < maxCols) {
          mapped.push({ value: null });
        }
        return mapped;
      });
      const exactRange = formatA1({
        sheet: addr.sheet,
        startCol: addr.startCol,
        startRow: addr.startRow,
        endCol: addr.startCol + maxCols - 1,
        endRow: addr.startRow + Math.max(cells.length - 1, 0),
      });
      return {
        kind: 'write_range',
        range: exactRange,
        cells,
        reason: args.reason as string,
      };
    }
    case 'clear_range':
      return { kind: 'clear_range', range: args.range as string, reason: args.reason as string };
    case 'format_range':
      return {
        kind: 'format_range',
        range: args.range as string,
        format: args.format as Extract<ChangeOp, { kind: 'format_range' }>['format'],
        reason: args.reason as string,
      };
    case 'add_sheet':
      return { kind: 'add_sheet', name: args.name as string, reason: args.reason as string };
    case 'rename_sheet':
      return {
        kind: 'rename_sheet',
        name: args.name as string,
        newName: args.newName as string,
        reason: args.reason as string,
      };
    case 'delete_sheet':
      return { kind: 'delete_sheet', name: args.name as string, reason: args.reason as string };
    case 'insert_rows':
    case 'delete_rows':
      return {
        kind: name,
        sheet: args.sheet as string,
        at: (args.at as number) - 1, // model speaks 1-based rows
        count: args.count as number,
        reason: args.reason as string,
      };
    case 'insert_cols':
    case 'delete_cols':
      return {
        kind: name,
        sheet: args.sheet as string,
        at: colToIndex((args.at as string).toUpperCase()),
        count: args.count as number,
        reason: args.reason as string,
      };
    case 'sort_range': {
      const range = args.range as string;
      const addr = parseA1(range);
      const keyOffset = colToIndex((args.keyColumn as string).toUpperCase()) - addr.startCol;
      if (keyOffset < 0 || keyOffset > addr.endCol - addr.startCol) {
        throw new Error(`keyColumn ${String(args.keyColumn)} is outside the range ${range}`);
      }
      return {
        kind: 'sort_range',
        range,
        keyOffset,
        ascending: args.ascending as boolean,
        hasHeader: args.hasHeader as boolean,
        reason: args.reason as string,
      };
    }
    case 'import_local_data':
      throw new Error('import_local_data is handled asynchronously in executeToolCall');
    default:
      throw new Error(`Not a write tool: ${name}`);
  }
}

function summarize(name: ToolName, args: Record<string, unknown>): string {
  switch (name) {
    case 'read_range':
    case 'write_range':
    case 'clear_range':
    case 'format_range':
    case 'sort_range':
      return String(args.range ?? '');
    case 'add_sheet':
    case 'rename_sheet':
    case 'delete_sheet':
      return String(args.name ?? '');
    case 'insert_rows':
    case 'delete_rows':
    case 'insert_cols':
    case 'delete_cols':
      return `${String(args.sheet ?? '')} @${String(args.at ?? '')}×${String(args.count ?? 1)}`;
    case 'find':
      return String(args.query ?? '');
    case 'import_local_data':
      return String(args.filePath ?? '');
    default:
      return '';
  }
}

/**
 * Early feedback for the model: the same sheet resolution staging performs,
 * run against the turn's cached map. Staging re-checks against a fresh map.
 */
async function findSheetProblem(
  ops: ChangeOp[],
  getCachedMap: () => Promise<WorkbookMap>,
): Promise<string | null> {
  let map: WorkbookMap;
  try {
    map = await getCachedMap();
  } catch {
    return null; // workbook unreachable; staging will report it
  }
  try {
    resolveSheetTargets(ops, map);
    return null;
  } catch (e) {
    return e instanceof Error ? e.message : String(e);
  }
}

function errorResult(message: string): string {
  return JSON.stringify({ ok: false, error: message });
}

function safeParseJson(s: string): unknown {
  try {
    return JSON.parse(s);
  } catch {
    return s;
  }
}

/**
 * Local models sometimes omit the sheet prefix ("A1:B5"), pass the workbook
 * filename as the sheet prefix ("Book 2026-10-28.xlsx!A1:B5"), or default to
 * "Sheet1!A1:B5" when the workbook tab has a different name. Resolves these to
 * the actual worksheet in the workbook.
 */
async function normalizeRangeForWorkbook(
  range: string,
  workbookName: string,
  getCachedMap: () => Promise<WorkbookMap>,
  pendingOps?: ChangeOp[],
): Promise<string> {
  const trimmed = range.trim();
  const bang = trimmed.lastIndexOf('!');
  if (bang < 0) {
    try {
      const map = await getCachedMap();
      const defaultSheet = map.sheets[0]?.name ?? 'Sheet1';
      const quoted = /^[A-Za-z0-9_]+$/.test(defaultSheet) ? defaultSheet : `'${defaultSheet}'`;
      return `${quoted}!${trimmed}`;
    } catch {
      return trimmed;
    }
  }

  const rawSheet = trimmed.slice(0, bang).replace(/^'|'$/g, '').trim();
  const addrPart = trimmed.slice(bang + 1);
  const sheetLower = rawSheet.toLowerCase();

  try {
    const map = await getCachedMap();
    const existing = map.sheets.find((s) => s.name.trim().toLowerCase() === sheetLower);
    if (existing) {
      const quoted = /^[A-Za-z0-9_]+$/.test(existing.name) ? existing.name : `'${existing.name}'`;
      return `${quoted}!${addrPart}`;
    }
    const addedInTurn = pendingOps?.some(
      (op) =>
        (op.kind === 'add_sheet' && op.name.trim().toLowerCase() === sheetLower) ||
        (op.kind === 'rename_sheet' && op.newName.trim().toLowerCase() === sheetLower),
    );
    if (addedInTurn) {
      return trimmed;
    }
    const wbLower = workbookName.trim().toLowerCase();
    const wbBaseLower = wbLower.replace(/\.(xlsx|xlsm|xlsb|xls|csv)$/i, '');
    const looksLikeWorkbookFile =
      sheetLower === wbLower ||
      (wbBaseLower.length > 0 && sheetLower === wbBaseLower) ||
      /\.(xlsx|xlsm|xlsb|xls|csv)$/i.test(rawSheet);

    if (
      map.sheets.length > 0 &&
      (looksLikeWorkbookFile || sheetLower === 'sheet1' || map.sheets.length === 1)
    ) {
      const actualSheet = map.sheets[0]!.name;
      const quoted = /^[A-Za-z0-9_]+$/.test(actualSheet) ? actualSheet : `'${actualSheet}'`;
      return `${quoted}!${addrPart}`;
    }
  } catch {
    return trimmed;
  }
  return trimmed;
}

function tryParseEmulatedCall(text: string): { tool: string; args: unknown } | null {
  const trimmed = text.trim();
  // Tolerate code fences around the JSON.
  const body = trimmed.startsWith('```')
    ? trimmed.replace(/^```(?:json)?\s*/i, '').replace(/```\s*$/, '')
    : trimmed;
  if (!body.startsWith('{')) return null;
  try {
    const obj = JSON.parse(body) as Record<string, unknown>;
    if (typeof obj.tool === 'string' && 'args' in obj) {
      return { tool: obj.tool, args: obj.args };
    }
    if (typeof obj.name === 'string' && ('arguments' in obj || 'args' in obj)) {
      return { tool: obj.name, args: obj.arguments ?? obj.args };
    }
    return null;
  } catch {
    return null;
  }
}

const DEFERRED_OR_FAKE_STAGE_RE =
  /\b(?:(?:kept|keeping) the (?:prior |previous )?staged version|previous version is already staged|won['’]t overwrite it|rather than overwriting|(?:i|i['’]ve|i have|we|we['’]ve)\s+(?:just\s+)?staged\b|(?:can|will|i['’]ll)\s+(?:stage|reorganize|rearrange|create|build|update)\b[\s\S]{0,140}\b(?:next|if you['’]d like|if you want))\b/i;

function stripUserContextHeaders(userMessage: string): string {
  return userMessage
    .replace(/\[Attached Document:\s*[^\n\]]+?\]\r?\n```[^\r\n]*\r?\n[\s\S]*?\r?\n```/g, '')
    .replace(/\[(?:Active Selection|📎 Selection):[^\]]*\]/gi, '')
    .replace(/Relevant saved context:\n[\s\S]*$/i, '')
    .trim();
}

/**
 * Returns true when the model replied with plain text (`toolCalls.length === 0`) and `pendingOps.length === 0`
 * on a turn where it either:
 * 1. Refused to overwrite a prior turn's staged change-set or deferred staging to "next",
 * 2. Hallucinated in text that it "staged" changes in this turn without calling any write tool, or
 * 3. Received a direct user confirmation ("yes", "sounds good. let's do it") after a prior turn and
 *    still emitted no write tool call.
 */
function shouldRetryDeferredOrUnstagedWrite(
  assistantText: string,
  userMessage: string,
  history: ChatMessage[],
): boolean {
  if (DEFERRED_OR_FAKE_STAGE_RE.test(assistantText)) {
    return true;
  }

  const cleanUser = stripUserContextHeaders(userMessage);
  if (
    history.length > 0 &&
    cleanUser.length > 0 &&
    cleanUser.length <= 80 &&
    /^(?:yes|yep|yeah|sure|ok|okay|sounds good|let['’]?s do it|do it|go ahead|please do|stage it|apply it|proceed)[.!?\s,;:-]*(?:let['’]?s do it|do it|please|go ahead|stage it|now)*[.!?\s]*$/i.test(
      cleanUser,
    )
  ) {
    return true;
  }

  if (
    /\b(?:reorganize|re-organize|rearrange|transpose|reformat|rebuild)\b/i.test(cleanUser) &&
    !assistantText.trim().endsWith('?')
  ) {
    return true;
  }

  return false;
}

function buildDeferredWriteRetryNudge(
  userMessage: string,
  assistantText: string,
  history: ChatMessage[],
): string {
  const recentText = [
    ...history.slice(-4).map((m) => m.content),
    userMessage,
    assistantText,
  ].join('\n');

  const wantsHorizontalColumns =
    /rows below|below the starting|rather than up\s*(?:and\s*)?down/i.test(recentText);

  const layoutDirective = wantsHorizontalColumns
    ? ' Arrange the overarching group headers across columns in Row 1, the individual item/column headers from left to right in Row 2, and the initial values in Row 3 so additional entries can be added in the rows below.'
    : '';

  return `[System note: You replied in plain text without calling write_range, so NOTHING was staged in this turn. Prior turns' staged change-sets never block overwriting, and you must NEVER defer staging to "next" or ask permission when the user already asked for the change. Immediately call write_range starting at "Sheet1!A1" RIGHT NOW to stage the requested table.${layoutDirective}]`;
}

/**
 * Builds a decisive recovery prompt when a local reasoning model exhausts tokens or ruminates
 * in `<think>` / `reasoning_content`:
 * 1. Uses `trimRuminatingReasoning` to preserve the initial extracted numbers/rows and strip
 *    trailing circular deliberation.
 * 2. Locks a concrete layout decision when the user requested a horizontal / side-by-side
 *    orientation so the model emits `write_range` immediately without further debate.
 */
function buildReasoningRecoveryNudge(
  thinkingText: string,
  userMessage: string,
  baseNudge: string,
): string {
  const cleanedDraft = trimRuminatingReasoning(thinkingText.trim(), 2000);
  const combined = `${userMessage}\n${cleanedDraft}`;
  const wantsHorizontalColumns =
    /rows below|below the starting|rather than up\s*(?:and\s*)?down/i.test(combined);
  const wantsLeftToRight =
    /left\s+to\s+right|side[\s-]+by[\s-]+side|horizontal|on the left[\s\S]{0,60}on the right/i.test(
      combined,
    );

  const layoutLock = wantsHorizontalColumns
    ? '\n\nLAYOUT DECISION LOCKED: Use a single write_range call starting at "Sheet1!A1" where categories/items are arranged as COLUMNS from left to right (overarching group headers in Row 1, individual column headers from left to right in Row 2, and initial values in Row 3 so new entries can be added in the rows below). Do NOT deliberate or consider any other layout — emit write_range immediately.'
    : wantsLeftToRight
      ? '\n\nLAYOUT DECISION LOCKED: Commit immediately to a single write_range call starting at "Sheet1!A1" that places the requested left-to-right sections across columns (either as column headers from left to right with values in the rows below, or as side-by-side column blocks separated by a null spacer column). Do NOT deliberate or compare alternative layouts — emit write_range immediately.'
      : '';

  if (!cleanedDraft) {
    return `${baseNudge}${layoutLock}`;
  }

  return `${baseNudge}${layoutLock}\n\nHere is your cleaned reasoning draft — use these exact values/rows and emit the tool call immediately without re-drafting:\n"""\n${cleanedDraft}\n"""`;
}

/**
 * If a local reasoning model drafted a complete JSON tool call (or write_range args block)
 * inside `<think>` / `reasoning_content` and then hit `max_tokens` before emitting `tool_calls`,
 * extract the last valid JSON block so the turn succeeds immediately without another 5-minute retry.
 */
function tryExtractToolCallFromReasoning(
  reasoning: string,
  iteration: number,
): { id: string; name: string; argsJson: string } | null {
  // 1. Check explicit <tool_call> tags or ```json ... ``` blocks via extractToolCalls
  const extracted = extractToolCalls(reasoning);
  if (extracted.toolCalls.length > 0) {
    return extracted.toolCalls[extracted.toolCalls.length - 1]!;
  }

  // 2. Check markdown code fences for direct write_range / format_range / read_range argument objects
  //    e.g. ```json\n{ "range": "Sheet1!A1", "cells": [...], "reason": "..." }\n```
  const codeBlockRegex = /```(?:json)?\s*([\s\S]*?)\s*```/gi;
  const candidates: { name: string; argsJson: string }[] = [];
  let match: RegExpExecArray | null;
  while ((match = codeBlockRegex.exec(reasoning)) !== null) {
    const block = match[1]?.trim();
    if (!block || !block.startsWith('{')) continue;
    try {
      const obj = JSON.parse(block) as Record<string, unknown>;
      if (typeof obj.range === 'string' && Array.isArray(obj.cells) && typeof obj.reason === 'string') {
        candidates.push({ name: 'write_range', argsJson: JSON.stringify(obj) });
      } else if (
        typeof obj.range === 'string' &&
        typeof obj.format === 'object' &&
        obj.format !== null &&
        typeof obj.reason === 'string'
      ) {
        candidates.push({ name: 'format_range', argsJson: JSON.stringify(obj) });
      }
    } catch {
      // ignore non-JSON or incomplete blocks
    }
  }

  if (candidates.length > 0) {
    const last = candidates[candidates.length - 1]!;
    return {
      id: `rescued_${iteration}`,
      name: last.name,
      argsJson: last.argsJson,
    };
  }
  return null;
}

/** A syntax-only correction applied to a staged formula; reported back, never silent. */
interface FormulaCorrection {
  cell: string;
  from: string;
  to: string;
  fix: string;
}

/** A reference that looks wrong but is left as written; the model must fix or confirm it. */
interface FormulaDiagnostic {
  cell: string;
  formula: string;
  issue: string;
  candidates: string[];
}

/**
 * Fixes only what is unambiguous syntax (the formula's meaning can't change):
 * a missing leading `=`, `;` argument separators from non-en-US locales, and
 * missing closing parentheses at the end.
 */
function fixFormulaSyntax(formula: string): { formula: string; fixes: string[] } {
  const fixes: string[] = [];
  let f = formula.trim();
  if (!f.startsWith('=')) {
    f = `=${f}`;
    fixes.push('added the leading "="');
  }
  let out = '';
  let open = 0;
  let inQuotes = false;
  let inArray = false; // `;` separates rows inside an array constant
  let sawSemicolon = false;
  for (const ch of f) {
    if (ch === '"') inQuotes = !inQuotes;
    else if (!inQuotes) {
      if (ch === '{') inArray = true;
      else if (ch === '}') inArray = false;
      else if (ch === '(') open++;
      else if (ch === ')' && open > 0) open--;
      else if (ch === ';' && !inArray) {
        out += ',';
        sawSemicolon = true;
        continue;
      }
    }
    out += ch;
  }
  if (sawSemicolon) fixes.push('replaced ";" argument separators with en-US ","');
  if (open > 0 && !inQuotes) {
    out += ')'.repeat(open);
    fixes.push(`closed ${open} unbalanced parenthes${open === 1 ? 'is' : 'es'}`);
  }
  return { formula: out, fixes };
}

/**
 * Normalizes `write_range` cells before validation and staging. It changes
 * shape and syntax only, never what a formula refers to:
 * 1. Unwraps accidental 3-D single-row array nesting (`[[["Cash", 401900]]]` -> `[["Cash", 401900]]`).
 * 2. Applies {@link fixFormulaSyntax} to `{ formula }` cells, reporting each correction.
 * 3. Flags references that point at blank or label cells of the block being
 *    written, in the patterns local models get wrong: a vertical SUM shifted
 *    off its number block by a miscounted spacer row (and refs to the
 *    subtotals it shifted), and value columns referring to the label column
 *    beside them. A formula whose references land on values is never flagged,
 *    so a total of an earlier block, another column, an absolute or a
 *    cross-sheet reference keeps its meaning.
 */
function repairWriteRangeCells(
  rangeStr: string,
  rawCells: unknown[],
): { cells: unknown[]; corrections: FormulaCorrection[]; diagnostics: FormulaDiagnostic[] } {
  let addr: ReturnType<typeof parseA1> | undefined;
  try {
    addr = parseA1(rangeStr);
  } catch {
    addr = undefined;
  }
  const cellName = (r: number, c: number): string =>
    addr
      ? formatA1({
          sheet: addr.sheet,
          startRow: addr.startRow + r,
          startCol: addr.startCol + c,
          endRow: addr.startRow + r,
          endCol: addr.startCol + c,
        })
      : `row ${r + 1} col ${c + 1}`;

  const corrections: FormulaCorrection[] = [];
  const diagnostics: FormulaDiagnostic[] = [];
  const rows = rawCells.map((row, r) => {
    let current = row;
    while (Array.isArray(current) && current.length === 1 && Array.isArray(current[0])) {
      current = current[0];
    }
    if (!Array.isArray(current)) return current;
    return current.map((cell, c) => {
      if (
        typeof cell === 'object' &&
        cell !== null &&
        'formula' in cell &&
        typeof (cell as { formula: unknown }).formula === 'string'
      ) {
        const from = (cell as { formula: string }).formula;
        const { formula, fixes } = fixFormulaSyntax(from);
        if (fixes.length > 0) corrections.push({ cell: cellName(r, c), from, to: formula, fix: fixes.join('; ') });
        return { ...cell, formula };
      }
      return cell;
    });
  });
  if (!addr) return { cells: rows, corrections, diagnostics };
  const start = addr;

  const baseRow = start.startRow + 1; // 1-based sheet row for rows[0]
  const maxCols = Math.max(0, ...rows.map((r) => (Array.isArray(r) ? r.length : 0)));

  const sumRangeRe =
    /^=\s*SUM\(\s*(\$?)([A-Za-z]{1,3})(\$?)(\d+)\s*:\s*(\$?)([A-Za-z]{1,3})(\$?)(\d+)\s*\)$/i;
  const singleRefRe = /^=\s*(\$?)([A-Za-z]{1,3})(\$?)(\d+)\s*$/i;
  // Plain cell refs: not part of a range, a function name or a sheet-qualified ref.
  const cellTokenRe = /(^|[^:A-Za-z0-9_!'])(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?!\s*:|\s*\(|[A-Za-z0-9_])/g;
  const flag = (r: number, c: number, formula: string, issue: string, candidate: string) => {
    if (candidate !== formula) diagnostics.push({ cell: cellName(r, c), formula, issue, candidates: [candidate] });
  };

  for (let c = 0; c < maxCols; c++) {
    const colLetter = indexToCol(start.startCol + c).toUpperCase();
    const getCellAt = (colIdx: number, sheetRow: number): unknown => {
      const idx = sheetRow - baseRow;
      if (idx < 0 || idx >= rows.length) return undefined;
      const r = rows[idx];
      return Array.isArray(r) ? r[colIdx] : undefined;
    };
    const getCell = (sheetRow: number): unknown => getCellAt(c, sheetRow);
    const isNumericCellAt = (colIdx: number, sheetRow: number): boolean =>
      typeof getCellAt(colIdx, sheetRow) === 'number';
    const isNumericCell = (sheetRow: number): boolean => isNumericCellAt(c, sheetRow);
    const isFormulaCellAt = (colIdx: number, sheetRow: number): boolean => {
      const v = getCellAt(colIdx, sheetRow);
      return (
        typeof v === 'object' &&
        v !== null &&
        'formula' in v &&
        typeof (v as { formula: unknown }).formula === 'string'
      );
    };
    const isFormulaCell = (sheetRow: number): boolean => isFormulaCellAt(c, sheetRow);
    const isEmptyCell = (sheetRow: number): boolean => {
      const v = getCell(sheetRow);
      return v === undefined || v === null || v === '';
    };
    /** `col` is another column of this block holding no value on `refRow`, where this column holds one. */
    const isLabelColRef = (col: string, refRow: number, actualRow: number): boolean => {
      const otherColOffset = colToIndex(col) - start.startCol;
      return (
        col !== colLetter &&
        otherColOffset >= 0 &&
        otherColOffset < maxCols &&
        refRow < actualRow &&
        !isNumericCellAt(otherColOffset, refRow) &&
        !isFormulaCellAt(otherColOffset, refRow) &&
        (isNumericCell(refRow) || isFormulaCell(refRow))
      );
    };

    // Where the model believed each flagged subtotal sat (key), and where it is (value).
    const mentalToActualSubtotalRow = new Map<number, number>();
    const passThroughSubtotalRow = new Map<number, number>();
    const summedDetailRows = new Set<number>();

    for (let r = 0; r < rows.length; r++) {
      const actualRow = baseRow + r;
      const rowArr = rows[r];
      if (!Array.isArray(rowArr) || !isFormulaCell(actualRow)) continue;
      const formula = (rowArr[c] as { formula: string }).formula;

      // Side-by-side tables (e.g. Assets in A:B, Liabilities in D:E): a total in
      // value column E that sums label column D.
      const sumMatch = sumRangeRe.exec(formula);
      const sumCol = sumMatch?.[2]!.toUpperCase();
      if (sumMatch && sumCol === sumMatch[6]!.toUpperCase() && sumCol !== colLetter) {
        const otherColOffset = colToIndex(sumCol!) - start.startCol;
        const sRow = parseInt(sumMatch[4]!, 10);
        const eRow = parseInt(sumMatch[8]!, 10);
        if (otherColOffset >= 0 && otherColOffset < maxCols && sRow <= eRow && eRow < actualRow) {
          const rowSpan = Array.from({ length: eRow - sRow + 1 }, (_, i) => sRow + i);
          const otherHasValue = rowSpan.some(
            (k) => isNumericCellAt(otherColOffset, k) || isFormulaCellAt(otherColOffset, k),
          );
          if (!otherHasValue && rowSpan.some((k) => isNumericCell(k))) {
            flag(
              r,
              c,
              formula,
              `sums ${sumCol}${sRow}:${sumCol}${eRow}, which holds only labels or blanks in this block, while column ${colLetter} holds the numbers on those rows`,
              `=SUM(${sumMatch[1]}${colLetter}${sumMatch[3]}${sRow}:${sumMatch[5]}${colLetter}${sumMatch[7]}${eRow})`,
            );
          }
        }
        continue;
      }
      const isPlainArithmetic = !formula.includes('"') && !/[A-Za-z_]\w*\s*\(/.test(formula);
      if (!sumMatch && isPlainArithmetic) {
        const labelRefs = [...formula.matchAll(cellTokenRe)].filter((m) =>
          isLabelColRef(m[3]!.toUpperCase(), parseInt(m[5]!, 10), actualRow),
        );
        if (labelRefs.length > 0) {
          flag(
            r,
            c,
            formula,
            `refers to ${labelRefs.map((m) => `${m[3]}${m[5]}`).join(', ')}, holding a label or blank in this block, while column ${colLetter} holds the value on that row`,
            formula.replace(cellTokenRe, (full, prefix: string, d1: string, col: string, d2: string, rowDigits: string) =>
              isLabelColRef(col.toUpperCase(), parseInt(rowDigits, 10), actualRow)
                ? `${prefix}${d1}${colLetter}${d2}${rowDigits}`
                : full,
            ),
          );
          continue;
        }
      }

      // A vertical SUM as long as the number block right above it, but shifted
      // onto a blank row: the signature of a miscounted spacer row.
      if (sumMatch) {
        const sRow = parseInt(sumMatch[4]!, 10);
        const eRow = parseInt(sumMatch[8]!, 10);
        const numEnd = actualRow - 1;
        let numStart = numEnd + 1;
        while (numStart - 1 >= baseRow && isNumericCell(numStart - 1)) numStart--;
        if (numStart > numEnd) continue;
        if (sRow === numStart && eRow === numEnd) {
          for (let k = numStart; k <= numEnd; k++) summedDetailRows.add(k);
        } else if (
          sRow <= eRow &&
          eRow < actualRow &&
          eRow - sRow === numEnd - numStart &&
          Math.abs(numEnd - eRow) <= 5 &&
          Array.from({ length: eRow - sRow + 1 }, (_, i) => sRow + i).some((k) => isEmptyCell(k))
        ) {
          flag(
            r,
            c,
            formula,
            `sums ${colLetter}${sRow}:${colLetter}${eRow}, which includes blank rows of this block; the ${numEnd - numStart + 1} numbers directly above are ${colLetter}${numStart}:${colLetter}${numEnd}`,
            `=SUM(${sumMatch[1]}${colLetter}${sumMatch[3]}${numStart}:${sumMatch[5]}${colLetter}${sumMatch[7]}${numEnd})`,
          );
          for (let k = numStart; k <= numEnd; k++) summedDetailRows.add(k);
          // The model thought the detail block ended at eRow, so it thought this subtotal row was at eRow + 1.
          mentalToActualSubtotalRow.set(eRow + 1, actualRow);
        }
        continue;
      }

      // Refs to where the model believed a flagged subtotal was (e.g. =B24 for TOTAL LIABILITIES).
      const shiftedTo = (refRow: number): number | undefined => {
        const mapped = mentalToActualSubtotalRow.get(refRow);
        return mapped !== undefined &&
          mapped < actualRow &&
          (isEmptyCell(refRow) || summedDetailRows.has(refRow)) &&
          isFormulaCell(mapped)
          ? mapped
          : undefined;
      };
      const singleMatch = singleRefRe.exec(formula);
      if (singleMatch && singleMatch[2]!.toUpperCase() === colLetter) {
        const refRow = parseInt(singleMatch[4]!, 10);
        const mapped = shiftedTo(refRow);
        if (mapped !== undefined) {
          flag(
            r,
            c,
            formula,
            `refers to ${colLetter}${refRow}, where a subtotal flagged above was expected to be; it is at ${colLetter}${mapped}`,
            `=${singleMatch[1]}${colLetter}${singleMatch[3]}${mapped}`,
          );
          passThroughSubtotalRow.set(mapped, actualRow);
        } else if (isFormulaCell(refRow) && refRow < actualRow) {
          passThroughSubtotalRow.set(refRow, actualRow);
        }
        continue;
      }
      if (mentalToActualSubtotalRow.size > 0) {
        const shifted = [...formula.matchAll(cellTokenRe)].filter(
          (m) => m[3]!.toUpperCase() === colLetter && shiftedTo(parseInt(m[5]!, 10)) !== undefined,
        );
        if (shifted.length > 0) {
          flag(
            r,
            c,
            formula,
            `refers to ${shifted.map((m) => `${m[3]}${m[5]}`).join(', ')}, where subtotals flagged above were expected to be`,
            formula.replace(cellTokenRe, (full, prefix: string, d1: string, col: string, d2: string, rowDigits: string) => {
              if (col.toUpperCase() !== colLetter) return full;
              const refRow = parseInt(rowDigits, 10);
              const target = shiftedTo(refRow) ?? passThroughSubtotalRow.get(refRow);
              return target !== undefined && target < actualRow ? `${prefix}${d1}${colLetter}${d2}${target}` : full;
            }),
          );
        }
      }
    }
  }

  return { cells: rows, corrections, diagnostics };
}

export { MAX_REPAIR_ATTEMPTS_PER_CALL };
