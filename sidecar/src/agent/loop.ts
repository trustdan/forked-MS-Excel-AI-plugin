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
  type WorkbookMap,
} from '@excelai/shared';
import type { ChatMessage, Provider, ToolSpec } from '../providers/types.js';
import { withRetry } from '../providers/index.js';
import { extractThinking, extractToolCalls } from '../providers/embedded.js';
import type { WorkbookExecutor } from '../workbook/executor.js';
import type { ChangeSetManager } from '../changeset/manager.js';
import { buildSystemPrompt, emulatedToolInstructions } from './prompts.js';
import {
  generateImportChangeOps,
  loadAndParseLocalFile,
  parseMarkdownTable,
  parseTabularContent,
} from '../importers/tabular.js';
import type { McpManager } from '../mcp/manager.js';

export type AgentEvent =
  | { kind: 'thinking_delta'; text: string }
  | { kind: 'text_delta'; text: string }
  | { kind: 'thinking'; text: string; durationMs?: number }
  | { kind: 'text'; text: string }
  | { kind: 'tool_use'; tool: string; summary: string }
  | { kind: 'changeset_staged'; changeSetId: string; preview: ChangePreviewItem[] }
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
  const duplicateAssistantIndices = new Set<number>();
  const duplicateToolCallIds = new Set<string>();
  const recoveryMessageIndices = new Set<number>();
  let forceTextReplyNext = false;
  let emittedAnything = false;
  let reasoningRecoveryCount = 0;
  const MAX_REASONING_RECOVERY_ATTEMPTS = 2;
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
          temperature: reasoningRecoveryCount > 0 ? 0.2 : 0,
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

    if (assistantText.trim()) emit({ kind: 'text', text: assistantText });

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

    if (toolCalls.length === 0) {
      if (!assistantText.trim()) {
        if (
          response.stopReason === 'max_tokens' &&
          reasoningRecoveryCount < MAX_REASONING_RECOVERY_ATTEMPTS &&
          iteration + 1 < maxIterations
        ) {
          reasoningRecoveryCount++;
          const reasoningTail = thinkingText.trim().slice(-2000);
          const nudgeContent = reasoningTail
            ? `${RECOVERY_NUDGE}\n\nHere is the end of your previous reasoning draft — use these exact values/rows and emit the tool call immediately without re-drafting:\n"""\n${reasoningTail}\n"""`
            : RECOVERY_NUDGE;
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
      if (prevCall) {
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
    const cs = changeSets.stage(pendingOps);
    const preview = await changeSets.preview(cs.id);
    emit({ kind: 'changeset_staged', changeSetId: cs.id, preview });
    result.stagedChangeSetId = cs.id;
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

  if (typeof rawArgs === 'object' && rawArgs !== null) {
    const rec = rawArgs as Record<string, unknown>;
    if (typeof rec.range === 'string' && rec.range.trim().length > 0) {
      rec.range = await normalizeRangeForWorkbook(rec.range, workbookName, getCachedMap);
    }
    if (toolName === 'write_range' && Array.isArray(rec.cells)) {
      rec.cells = repairWriteRangeCells(
        typeof rec.range === 'string' ? rec.range : 'Sheet1!A1',
        rec.cells as unknown[],
      );
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

        const resolvedDiskPath = resolve(filePath);
        if (!existsSync(resolvedDiskPath)) {
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

          return errorResult(
            `File not found on local disk: "${filePath}". Do not guess directory paths (such as /mnt/data/) or retry import_local_data with guessed paths. If the user attached a document in chat, its full contents are already in the user message under [Attached Document: ...] — use write_range directly.`,
          );
        }

        const { rows, totalRows, totalCols, format } = loadAndParseLocalFile(filePath);
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
          note: `Staged ${rowsStaged}×${colsStaged} table at ${op.range} into the pending change-set${replacedPriorDraft ? ` (replaced prior draft at ${startCellA1})` : ''}. All formulas are balanced and valid. Do NOT call write_range on ${startCellA1} again — reply to the user with a concise summary to finish (or call format_range once if formatting is needed).`,
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

    if (
      msg.role === 'assistant' &&
      i < priorHistoryEndIdx &&
      msg.toolCalls &&
      msg.toolCalls.length > 0
    ) {
      let changed = false;
      const compactedCalls = msg.toolCalls.map((tc) => {
        if (tc.name !== 'write_range' || tc.argsJson.length <= 1200) return tc;
        try {
          const parsed = JSON.parse(tc.argsJson) as Record<string, unknown>;
          if (!Array.isArray(parsed.cells) || parsed.cells.length <= 8) return tc;
          const totalRows = parsed.cells.length;
          const omitted = totalRows - 6;
          const compactCells = [
            ...parsed.cells.slice(0, 3),
            [`... (${omitted} middle rows omitted from prior turn history; total ${totalRows} rows)`],
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
      prepared[i] = changed ? { ...msg, toolCalls: compactedCalls } : msg;
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
 * Local models sometimes omit the sheet prefix ("A1:B5") or pass the workbook
 * filename as the sheet prefix ("Book 2026-10-28.xlsx!A1:B5"). Resolves both
 * to the first actual worksheet in the workbook.
 */
async function normalizeRangeForWorkbook(
  range: string,
  workbookName: string,
  getCachedMap: () => Promise<WorkbookMap>,
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
  const wbLower = workbookName.trim().toLowerCase();
  const wbBaseLower = wbLower.replace(/\.(xlsx|xlsm|xlsb|xls|csv)$/i, '');
  const sheetLower = rawSheet.toLowerCase();
  const looksLikeWorkbookFile =
    sheetLower === wbLower ||
    (wbBaseLower.length > 0 && sheetLower === wbBaseLower) ||
    /\.(xlsx|xlsm|xlsb|xls|csv)$/i.test(rawSheet);

  if (!looksLikeWorkbookFile) return trimmed;

  try {
    const map = await getCachedMap();
    if (map.sheets.some((s) => s.name.toLowerCase() === sheetLower)) {
      return trimmed;
    }
    const actualSheet = map.sheets[0]?.name ?? 'Sheet1';
    const quoted = /^[A-Za-z0-9_]+$/.test(actualSheet) ? actualSheet : `'${actualSheet}'`;
    return `${quoted}!${addrPart}`;
  } catch {
    return trimmed;
  }
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

/**
 * Auto-balances missing closing parentheses in formula strings emitted by
 * local models (e.g. `"=SUM(B7:B10"` -> `"=SUM(B7:B10)"`).
 */
function balanceFormulaParens(formula: string): string {
  const trimmed = formula.trim();
  const withEq = trimmed.startsWith('=') ? trimmed : `=${trimmed}`;
  let open = 0;
  let inQuotes = false;
  for (let i = 0; i < withEq.length; i++) {
    const ch = withEq[i];
    if (ch === '"') {
      inQuotes = !inQuotes;
    } else if (!inQuotes) {
      if (ch === '(') open++;
      else if (ch === ')' && open > 0) open--;
    }
  }
  return open > 0 ? `${withEq}${')'.repeat(open)}` : withEq;
}

/**
 * Normalizes `write_range` cells before validation and staging:
 * 1. Unwraps accidental 3-D single-row array nesting (`[[["Cash", 401900]]]` -> `[["Cash", 401900]]`).
 * 2. Auto-balances missing closing parentheses on `{ formula }` cells (`"=SUM(B7:B10"` -> `"=SUM(B7:B10)"`).
 * 3. Repairs off-by-one vertical `SUM(ColA:ColB)` and downstream subtotal references caused when
 *    models miscount blank spacer rows (`[""]`) inside a multi-section statement.
 */
function repairWriteRangeCells(rangeStr: string, rawCells: unknown[]): unknown[] {
  const rows = rawCells.map((row) => {
    let current = row;
    while (Array.isArray(current) && current.length === 1 && Array.isArray(current[0])) {
      current = current[0];
    }
    if (!Array.isArray(current)) return current;
    return current.map((cell) => {
      if (
        typeof cell === 'object' &&
        cell !== null &&
        'formula' in cell &&
        typeof (cell as { formula: unknown }).formula === 'string'
      ) {
        return {
          ...cell,
          formula: balanceFormulaParens((cell as { formula: string }).formula),
        };
      }
      return cell;
    });
  });

  let addr;
  try {
    addr = parseA1(rangeStr);
  } catch {
    return rows;
  }

  const baseRow = addr.startRow + 1; // 1-based sheet row for rows[0]
  const maxCols = Math.max(0, ...rows.map((r) => (Array.isArray(r) ? r.length : 0)));

  const sumRangeRe =
    /^=\s*SUM\(\s*(\$?)([A-Za-z]{1,3})(\$?)(\d+)\s*:\s*(\$?)([A-Za-z]{1,3})(\$?)(\d+)\s*\)$/i;
  const singleRefRe = /^=\s*(\$?)([A-Za-z]{1,3})(\$?)(\d+)\s*$/i;

  for (let c = 0; c < maxCols; c++) {
    const colLetter = indexToCol(addr.startCol + c).toUpperCase();
    const getCell = (sheetRow: number): unknown => {
      const idx = sheetRow - baseRow;
      if (idx < 0 || idx >= rows.length) return undefined;
      const r = rows[idx];
      return Array.isArray(r) ? r[c] : undefined;
    };
    const isNumericCell = (sheetRow: number): boolean => typeof getCell(sheetRow) === 'number';
    const isFormulaCell = (sheetRow: number): boolean => {
      const v = getCell(sheetRow);
      return (
        typeof v === 'object' &&
        v !== null &&
        'formula' in v &&
        typeof (v as { formula: unknown }).formula === 'string'
      );
    };
    const isEmptyCell = (sheetRow: number): boolean => {
      const v = getCell(sheetRow);
      return v === undefined || v === null || v === '';
    };

    const mentalToActualSubtotalRow = new Map<number, number>();
    const passThroughSubtotalRow = new Map<number, number>();
    const summedDetailRows = new Set<number>();

    for (let r = 0; r < rows.length; r++) {
      const actualRow = baseRow + r;
      const rowArr = rows[r];
      if (!Array.isArray(rowArr) || !isFormulaCell(actualRow)) continue;
      const formulaObj = rowArr[c] as { formula: string };
      const formula = formulaObj.formula;

      // 1. Check vertical =SUM(Col<sRow>:Col<eRow>)
      const sumMatch = sumRangeRe.exec(formula);
      if (
        sumMatch &&
        sumMatch[2]!.toUpperCase() === colLetter &&
        sumMatch[6]!.toUpperCase() === colLetter
      ) {
        const sRow = parseInt(sumMatch[4]!, 10);
        const eRow = parseInt(sumMatch[8]!, 10);

        const numEnd = actualRow - 1;
        let numStart = numEnd + 1;
        while (numStart - 1 >= baseRow && isNumericCell(numStart - 1)) {
          numStart--;
        }

        if (numStart <= numEnd) {
          if (sRow === numStart && eRow === numEnd) {
            for (let k = numStart; k <= numEnd; k++) summedDetailRows.add(k);
          } else if (
            sRow <= eRow &&
            eRow < actualRow &&
            eRow - sRow === numEnd - numStart &&
            Math.abs(numEnd - eRow) <= 5 &&
            (eRow < numEnd ||
              Array.from({ length: eRow - sRow + 1 }, (_, i) => sRow + i).some((k) =>
                isEmptyCell(k),
              ))
          ) {
            rowArr[c] = {
              formula: `=SUM(${sumMatch[1]}${colLetter}${sumMatch[3]}${numStart}:${sumMatch[5]}${colLetter}${sumMatch[7]}${numEnd})`,
            };
            for (let k = numStart; k <= numEnd; k++) summedDetailRows.add(k);
            // The model thought the detail block ended at eRow, so it thought this subtotal row was at eRow + 1
            mentalToActualSubtotalRow.set(eRow + 1, actualRow);
          }
        }
        continue;
      }

      // 2. Check single-cell pass-through subtotal =Col<refRow> (e.g. =B24 on TOTAL LIABILITIES)
      const singleMatch = singleRefRe.exec(formula);
      if (singleMatch && singleMatch[2]!.toUpperCase() === colLetter) {
        const refRow = parseInt(singleMatch[4]!, 10);
        const mappedRow = mentalToActualSubtotalRow.get(refRow);
        if (
          mappedRow !== undefined &&
          mappedRow < actualRow &&
          (isEmptyCell(refRow) || summedDetailRows.has(refRow)) &&
          isFormulaCell(mappedRow)
        ) {
          rowArr[c] = {
            formula: `=${singleMatch[1]}${colLetter}${singleMatch[3]}${mappedRow}`,
          };
          passThroughSubtotalRow.set(mappedRow, actualRow);
        } else if (isFormulaCell(refRow) && refRow < actualRow) {
          passThroughSubtotalRow.set(refRow, actualRow);
        }
        continue;
      }

      // 3. Check multi-cell subtotal formulas (e.g. =B25+B30) when earlier subtotals shifted
      if (mentalToActualSubtotalRow.size > 0) {
        const cellTokenRe = /(^|[^:A-Za-z0-9_])(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?!\s*:)/g;
        let hasShiftedSubtotalRef = false;
        let m: RegExpExecArray | null;
        while ((m = cellTokenRe.exec(formula)) !== null) {
          if (m[3]!.toUpperCase() !== colLetter) continue;
          const refRow = parseInt(m[5]!, 10);
          const mapped = mentalToActualSubtotalRow.get(refRow);
          if (
            mapped !== undefined &&
            mapped < actualRow &&
            (isEmptyCell(refRow) || summedDetailRows.has(refRow)) &&
            isFormulaCell(mapped)
          ) {
            hasShiftedSubtotalRef = true;
            break;
          }
        }

        if (hasShiftedSubtotalRef) {
          const repaired = formula.replace(
            /(^|[^:A-Za-z0-9_])(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?!\s*:)/g,
            (full, prefix: string, d1: string, col: string, d2: string, rowDigits: string) => {
              if (col.toUpperCase() !== colLetter) return full;
              const refRow = parseInt(rowDigits, 10);
              const mapped = mentalToActualSubtotalRow.get(refRow);
              if (
                mapped !== undefined &&
                mapped < actualRow &&
                (isEmptyCell(refRow) || summedDetailRows.has(refRow)) &&
                isFormulaCell(mapped)
              ) {
                return `${prefix}${d1}${colLetter}${d2}${mapped}`;
              }
              const passThrough = passThroughSubtotalRow.get(refRow);
              if (passThrough !== undefined && passThrough < actualRow) {
                return `${prefix}${d1}${colLetter}${d2}${passThrough}`;
              }
              return full;
            },
          );
          rowArr[c] = { formula: repaired };
        }
      }
    }
  }

  return rows;
}

export { MAX_REPAIR_ATTEMPTS_PER_CALL };
