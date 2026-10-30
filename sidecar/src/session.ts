import { makeEnvelope } from '@excelai/shared';
import type { WebSocket } from 'ws';
import { runAgentTurn } from './agent/loop.js';
import { ChangeSetManager } from './changeset/manager.js';
import type { WorkbookExecutor } from './workbook/executor.js';
import type { ChatMessage, Provider } from './providers/types.js';
import type { Ledger } from './ledger.js';
import type { MemoryStore } from './memory/store.js';
import type { McpManager } from './mcp/manager.js';

const MAX_UNRECOVERED_SHOWN = 20;

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
    const existing = this.changeSets.get(changeSetId);
    if (!this.known(changeSetId)) return;
    // A duplicate approval: the first handler reports the outcome.
    if (existing?.status === 'applying') return;
    try {
      if (existing?.status !== 'applied') {
        const result = await this.changeSets.apply(changeSetId);
        if (!result.ok) {
          this.emit({
            kind: 'error',
            message: this.describeFailure(
              changeSetId,
              result.verification === 'incomplete'
                ? `Change-set could not be verified, so it was rolled back:\n${result.problems.join('\n')}`
                : `Change-set verification failed:\n${result.problems.join('\n')}`,
            ),
          });
        }
      }
    } catch (e) {
      this.emit({ kind: 'error', message: this.describeFailure(changeSetId, e) });
    } finally {
      this.emitStatus(changeSetId);
    }
  }

  /** Reports an id this sidecar does not hold (e.g. staged before a restart) as `unknown`. */
  private known(changeSetId: string): boolean {
    if (this.changeSets.get(changeSetId)) return true;
    this.emit({
      kind: 'error',
      message: 'That change-set is no longer available (the sidecar may have restarted). Ask again to re-propose it.',
    });
    this.emitStatus(changeSetId);
    return false;
  }

  /** Tells the add-in where a change-set ended up, so its card matches the sidecar. */
  private emitStatus(changeSetId: string): void {
    const status = this.changeSets.get(changeSetId)?.status ?? 'unknown';
    if (status === 'applying' || status === 'undoing') return;
    this.emit({ kind: 'changeset_status', changeSetId, status });
  }

  private describeFailure(changeSetId: string, e: unknown): string {
    const message = e instanceof Error ? e.message : String(e);
    const cs = this.changeSets.get(changeSetId);
    if (cs?.status === 'rolled_back') {
      return `${message}\nThe change-set was rolled back; the workbook is unchanged.`;
    }
    if (cs?.status !== 'failed_partial') return message;
    if (!cs.unrecovered?.length) {
      return `Change-set failed partway; the workbook may be partially changed. ${message}`;
    }
    const shown = cs.unrecovered.slice(0, MAX_UNRECOVERED_SHOWN).join(', ');
    const more = cs.unrecovered.length - MAX_UNRECOVERED_SHOWN;
    return (
      `${message}\nThe workbook may be partially changed; these were not restored: ${shown}` +
      (more > 0 ? ` and ${more} more` : '')
    );
  }

  handleReject(changeSetId: string): void {
    if (!this.known(changeSetId)) return;
    try {
      this.changeSets.reject(changeSetId);
    } catch (e) {
      this.emit({ kind: 'error', message: e instanceof Error ? e.message : String(e) });
    } finally {
      this.emitStatus(changeSetId);
    }
  }

  async handleUndo(changeSetId: string): Promise<void> {
    if (!this.known(changeSetId)) return;
    if (this.changeSets.get(changeSetId)?.status === 'undoing') return;
    try {
      await this.changeSets.undo(changeSetId);
    } catch (e) {
      this.emit({ kind: 'error', message: this.describeFailure(changeSetId, e) });
    } finally {
      this.emitStatus(changeSetId);
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
