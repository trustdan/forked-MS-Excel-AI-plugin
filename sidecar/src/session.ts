import { makeEnvelope } from '@excelai/shared';
import type { WebSocket } from 'ws';
import { runAgentTurn } from './agent/loop.js';
import { ChangeSetManager } from './changeset/manager.js';
import type { WorkbookExecutor } from './workbook/executor.js';
import type { ChatMessage, Provider } from './providers/types.js';
import type { Ledger } from './ledger.js';
import type { MemoryStore } from './memory/store.js';
import type { McpManager } from './mcp/manager.js';

/**
 * One session per connected workbook. Sessions are scoped by workbookId so
 * multiple open workbooks (or Excel instances) never share state.
 */
export class Session {
  readonly changeSets: ChangeSetManager;
  private history: ChatMessage[] = [];
  private busy = false;
  private currentAbortController: AbortController | null = null;

  constructor(
    private ws: WebSocket,
    readonly workbookId: string,
    private workbookName: string,
    private executor: WorkbookExecutor,
    private readonly getProvider: () => Provider,
    private readonly ledger: Ledger,
    private readonly memory: MemoryStore,
    private readonly mcpManager?: McpManager,
  ) {
    this.changeSets = new ChangeSetManager(executor, workbookId);
  }

  /**
   * Rebinds the session to a new WebSocket connection and BridgeExecutor when
   * the add-in reconnects, preserving staged change-sets and chat history.
   */
  rebind(ws: WebSocket, workbookName: string, executor: WorkbookExecutor): void {
    this.ws = ws;
    this.workbookName = workbookName;
    this.executor = executor;
    this.changeSets.setExecutor(executor);
  }

  isCurrentSocket(ws: WebSocket): boolean {
    return this.ws === ws;
  }

  private emit(payload: Parameters<typeof makeEnvelope<'agent_event', unknown>>[1]): void {
    if (this.ws.readyState === this.ws.OPEN) {
      this.ws.send(JSON.stringify(makeEnvelope('agent_event', payload)));
    }
  }

  handleAbort(reason?: string): void {
    if (this.currentAbortController) {
      console.log(`[chat] aborting in-flight turn: ${reason ?? 'user requested'}`);
      this.currentAbortController.abort();
      this.executor.abortAll?.(reason ?? 'Turn aborted');
    }
  }

  async handleChat(message: string): Promise<void> {
    if (this.busy) {
      this.emit({ kind: 'error', message: 'The agent is already working on a request.' });
      return;
    }
    this.busy = true;
    const abortController = new AbortController();
    this.currentAbortController = abortController;
    const startedAt = Date.now();
    let totalIn = 0;
    let totalOut = 0;
    try {
      const provider = this.getProvider();

      // Inject relevant memories (user + this workbook) into the turn.
      const memories = this.memory.search(message, ['user', this.workbookId], 5);
      const memoryNote =
        memories.length > 0
          ? `\n\nRelevant saved context:\n${memories.map((m) => `- ${m.content}`).join('\n')}`
          : '';

      const result = await runAgentTurn({
        provider,
        executor: this.executor,
        changeSets: this.changeSets,
        workbookName: this.workbookName,
        history: this.history,
        userMessage: message + memoryNote,
        abortSignal: abortController.signal,
        mcpManager: this.mcpManager,
        onEvent: (ev) => this.emit(ev),
        onUsage: (usage) => {
          totalIn += usage.inputTokens;
          totalOut += usage.outputTokens;
          this.ledger.record({
            provider: provider.id,
            model: provider.model,
            feature: 'chat',
            ...usage,
          });
        },
      });
      if (!abortController.signal.aborted) {
        this.history = trimHistory(result.history, 24);
      }
      console.log(
        `[chat] ${this.workbookName} · ${provider.model} · ${Date.now() - startedAt}ms · tokens ${totalIn} in / ${totalOut} out${result.stagedChangeSetId ? ' · change-set staged' : ''}${abortController.signal.aborted ? ' · aborted' : ''}`,
      );
    } catch (e) {
      if (abortController.signal.aborted) {
        console.log(`[chat] turn aborted after ${Date.now() - startedAt}ms`);
      } else {
        const message = e instanceof Error ? e.message : String(e);
        console.error(`[chat] turn failed after ${Date.now() - startedAt}ms: ${message}`);
        this.emit({ kind: 'error', message });
      }
    } finally {
      this.busy = false;
      this.currentAbortController = null;
      this.emit({ kind: 'done' });
    }
  }

  async handleApprove(changeSetId: string): Promise<void> {
    try {
      const existing = this.changeSets.get(changeSetId);
      if (existing?.status === 'applied') {
        this.emit({ kind: 'changeset_applied', changeSetId });
        return;
      }
      const result = await this.changeSets.apply(changeSetId);
      if (result.ok) {
        this.emit({ kind: 'changeset_applied', changeSetId });
      } else {
        this.emit({
          kind: 'error',
          message: `Change-set verification failed and was rolled back:\n${result.problems.join('\n')}`,
        });
      }
    } catch (e) {
      this.emit({ kind: 'error', message: e instanceof Error ? e.message : String(e) });
    }
  }

  handleReject(changeSetId: string): void {
    try {
      this.changeSets.reject(changeSetId);
    } catch (e) {
      this.emit({ kind: 'error', message: e instanceof Error ? e.message : String(e) });
    }
  }

  async handleUndo(changeSetId: string): Promise<void> {
    try {
      await this.changeSets.undo(changeSetId);
      this.emit({ kind: 'changeset_undone', changeSetId });
    } catch (e) {
      this.emit({ kind: 'error', message: e instanceof Error ? e.message : String(e) });
    }
  }
}

/**
 * Bounds conversation history while ensuring it always starts at a `user`
 * message boundary so `tool` messages are never orphaned from their `assistant` call.
 */
function trimHistory(history: ChatMessage[], maxMessages: number): ChatMessage[] {
  if (history.length <= maxMessages) return history;
  let start = history.length - maxMessages;
  while (start < history.length && history[start]?.role !== 'user') {
    start++;
  }
  return start < history.length ? history.slice(start) : history.slice(-maxMessages);
}
