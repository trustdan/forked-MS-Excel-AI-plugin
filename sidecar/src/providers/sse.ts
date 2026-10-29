import { ProviderError, type ChatDelta } from './types.js';

/**
 * Maximum cumulative bytes read from a single SSE stream before aborting
 * to prevent memory exhaustion on runaway local servers (`structural-integrity-manager` invariant).
 */
export const MAX_SSE_STREAM_BYTES = 8 * 1024 * 1024; // 8 MiB

/**
 * Maximum length of a single buffered SSE line before rejecting as malformed.
 */
export const MAX_SSE_LINE_CHARS = 4 * 1024 * 1024; // 4 MiB

/**
 * Maximum accumulated characters for content, reasoning, or tool argument buffers.
 */
export const MAX_ACCUMULATED_TEXT_CHARS = 4 * 1024 * 1024; // 4 MiB

/**
 * Maximum time to wait for the next SSE chunk before aborting a stalled stream.
 */
export const DEFAULT_SSE_INACTIVITY_TIMEOUT_MS = 180_000; // 3 minutes

export interface SseEvent {
  event?: string;
  data: string;
}

/**
 * Reads an HTTP response body as a Server-Sent Events (SSE) stream with strict
 * memory bounds and deterministic `AbortSignal` reader cleanup.
 */
export async function readSseStream(
  res: Response,
  onEvent: (event: SseEvent) => void,
  options?: {
    abortSignal?: AbortSignal;
    maxStreamBytes?: number;
    maxLineChars?: number;
    inactivityTimeoutMs?: number;
  },
): Promise<void> {
  const body = res.body;
  if (!body) {
    throw new ProviderError('SSE response has no readable body');
  }

  const maxStreamBytes = options?.maxStreamBytes ?? MAX_SSE_STREAM_BYTES;
  const maxLineChars = options?.maxLineChars ?? MAX_SSE_LINE_CHARS;
  const inactivityTimeoutMs = options?.inactivityTimeoutMs ?? DEFAULT_SSE_INACTIVITY_TIMEOUT_MS;
  const abortSignal = options?.abortSignal;

  if (abortSignal?.aborted) {
    throw createAbortError();
  }

  const reader = body.getReader();
  const decoder = new TextDecoder();
  let totalBytes = 0;
  let buffer = '';
  let currentEvent: string | undefined;
  let dataLines: string[] = [];

  const onAbort = (): void => {
    void reader.cancel().catch(() => {});
  };
  abortSignal?.addEventListener('abort', onAbort, { once: true });

  const readNextChunk = async (): Promise<ReadableStreamReadResult<Uint8Array>> => {
    if (inactivityTimeoutMs <= 0) {
      return reader.read();
    }
    let timer: ReturnType<typeof setTimeout> | undefined;
    try {
      return await Promise.race([
        reader.read(),
        new Promise<never>((_, reject) => {
          timer = setTimeout(() => {
            void reader.cancel().catch(() => {});
            reject(
              new ProviderError(
                `SSE stream timed out after ${Math.round(inactivityTimeoutMs / 1000)}s of inactivity. Check if your local model server is stuck or out of context memory.`,
                504,
                false,
              ),
            );
          }, inactivityTimeoutMs);
        }),
      ]);
    } finally {
      if (timer !== undefined) clearTimeout(timer);
    }
  };

  const flushBlock = (): boolean => {
    if (dataLines.length === 0) {
      currentEvent = undefined;
      return false;
    }
    const data = dataLines.join('\n');
    const ev = currentEvent;
    dataLines = [];
    currentEvent = undefined;
    if (data.trim() === '[DONE]') {
      return true;
    }
    onEvent({ ...(ev !== undefined ? { event: ev } : {}), data });
    return false;
  };

  try {
    while (true) {
      if (abortSignal?.aborted) {
        throw createAbortError();
      }
      const { value, done } = await readNextChunk();
      if (abortSignal?.aborted) {
        throw createAbortError();
      }
      if (done) {
        buffer += decoder.decode();
        break;
      }
      if (value) {
        totalBytes += value.byteLength;
        if (totalBytes > maxStreamBytes) {
          throw new ProviderError(
            `SSE stream exceeded maximum allowed size (${maxStreamBytes} bytes)`,
            413,
            false,
          );
        }
        buffer += decoder.decode(value, { stream: true });
      }

      if (buffer.length > maxLineChars * 2) {
        throw new ProviderError(
          `SSE line buffer exceeded maximum allowed length (${maxLineChars} chars)`,
          413,
          false,
        );
      }

      let newlineIdx: number;
      while ((newlineIdx = buffer.indexOf('\n')) !== -1) {
        let line = buffer.slice(0, newlineIdx);
        buffer = buffer.slice(newlineIdx + 1);
        if (line.endsWith('\r')) {
          line = line.slice(0, -1);
        }
        if (line.length > maxLineChars) {
          throw new ProviderError(
            `SSE line exceeded maximum allowed length (${maxLineChars} chars)`,
            413,
            false,
          );
        }
        if (line === '') {
          const isDone = flushBlock();
          if (isDone) return;
          continue;
        }
        if (line.startsWith(':')) {
          continue;
        }
        if (line.startsWith('event:')) {
          currentEvent = line.slice(6).trim();
        } else if (line.startsWith('data:')) {
          const raw = line.slice(5);
          dataLines.push(raw.startsWith(' ') ? raw.slice(1) : raw);
        }
      }
    }

    // Flush any remaining trailing lines if the stream closed without a trailing blank line
    if (buffer.length > 0) {
      const remainingLines = buffer.split(/\r?\n/);
      for (const line of remainingLines) {
        if (line === '') {
          const isDone = flushBlock();
          if (isDone) return;
        } else if (line.startsWith('event:')) {
          currentEvent = line.slice(6).trim();
        } else if (line.startsWith('data:')) {
          const raw = line.slice(5);
          dataLines.push(raw.startsWith(' ') ? raw.slice(1) : raw);
        }
      }
    }
    flushBlock();
  } finally {
    abortSignal?.removeEventListener('abort', onAbort);
    try {
      await reader.cancel();
    } catch {
      // ignore cancel errors on already-closed streams
    }
    try {
      reader.releaseLock();
    } catch {
      // ignore
    }
  }
}

function createAbortError(): Error {
  const err = new Error('AbortError: Generation canceled by user');
  err.name = 'AbortError';
  return err;
}

const OPEN_THINK_TAGS = ['<think>', '<thinking>'];
const CLOSE_THINK_TAGS = ['</think>', '</thinking>'];
const TOOL_CALL_OPEN_TAG = '<tool_call>';

/**
 * Stateful incremental splitter that routes incoming content deltas into
 * `reasoning` (when inside `<think>...</think>` or `<thinking>...</thinking>`)
 * vs `text` (normal assistant reply), while suppressing raw `<tool_call>` blocks
 * from leaking into live UI text deltas.
 */
export class ThinkStreamSplitter {
  private mode: 'text' | 'thinking' | 'tool_call' = 'text';
  private pending = '';
  private emittedAnyText = false;

  constructor(private readonly onDelta?: (delta: ChatDelta) => void) {}

  /**
   * Push an explicit reasoning delta (e.g. `delta.reasoning_content` or Anthropic `thinking_delta`).
   */
  pushReasoning(chunk: string): void {
    if (!chunk) return;
    this.onDelta?.({ reasoning: chunk });
  }

  /**
   * Push a raw content chunk that may contain inline `<think>`, `</think>`, or `<tool_call>` tags.
   */
  pushContent(chunk: string): void {
    if (!chunk || !this.onDelta) return;
    this.pending += chunk;
    this.drain(false);
  }

  /**
   * Flush any remaining partial buffer when the stream finishes.
   */
  flush(): void {
    if (!this.onDelta || !this.pending) return;
    this.drain(true);
  }

  private emitText(str: string): void {
    if (!str) return;
    if (!this.emittedAnyText) {
      const trimmedStart = str.replace(/^\s+/, '');
      if (!trimmedStart) return;
      this.emittedAnyText = true;
      this.onDelta?.({ text: trimmedStart });
      return;
    }
    this.onDelta?.({ text: str });
  }

  private drain(final: boolean): void {
    while (this.pending.length > 0) {
      if (this.mode === 'tool_call') {
        // Once a raw <tool_call> tag starts in content, suppress subsequent content
        // from live text_delta so raw tool JSON/XML never flashes in the chat bubble.
        this.pending = '';
        return;
      }

      if (this.mode === 'text') {
        const lower = this.pending.toLowerCase();

        // Check for earliest opening <think>, <thinking>, or <tool_call>
        const match = findEarliestTag(lower, [...OPEN_THINK_TAGS, TOOL_CALL_OPEN_TAG]);
        if (match) {
          if (match.index > 0) {
            this.emitText(this.pending.slice(0, match.index));
          }
          this.pending = this.pending.slice(match.index + match.tag.length);
          this.mode = match.tag === TOOL_CALL_OPEN_TAG ? 'tool_call' : 'thinking';
          continue;
        }

        if (final) {
          this.emitText(this.pending);
          this.pending = '';
          return;
        }

        // Hold back any trailing prefix of `<think>`, `<thinking>`, or `<tool_call>`
        const holdback = computeTrailingTagPrefixLength(lower, [
          ...OPEN_THINK_TAGS,
          TOOL_CALL_OPEN_TAG,
        ]);
        const safeLen = this.pending.length - holdback;
        if (safeLen > 0) {
          this.emitText(this.pending.slice(0, safeLen));
          this.pending = this.pending.slice(safeLen);
        }
        return;
      }

      if (this.mode === 'thinking') {
        const lower = this.pending.toLowerCase();
        const match = findEarliestTag(lower, [...CLOSE_THINK_TAGS, TOOL_CALL_OPEN_TAG]);
        if (match) {
          if (match.index > 0) {
            this.pushReasoning(this.pending.slice(0, match.index));
          }
          this.pending = this.pending.slice(match.index + match.tag.length);
          this.mode = match.tag === TOOL_CALL_OPEN_TAG ? 'tool_call' : 'text';
          continue;
        }

        if (final) {
          this.pushReasoning(this.pending);
          this.pending = '';
          return;
        }

        const holdback = computeTrailingTagPrefixLength(lower, [
          ...CLOSE_THINK_TAGS,
          TOOL_CALL_OPEN_TAG,
        ]);
        const safeLen = this.pending.length - holdback;
        if (safeLen > 0) {
          this.pushReasoning(this.pending.slice(0, safeLen));
          this.pending = this.pending.slice(safeLen);
        }
        return;
      }
    }
  }
}

function findEarliestTag(
  lowerText: string,
  tags: string[],
): { index: number; tag: string } | null {
  let bestIndex = -1;
  let bestTag = '';
  for (const tag of tags) {
    const idx = lowerText.indexOf(tag);
    if (idx !== -1 && (bestIndex === -1 || idx < bestIndex || (idx === bestIndex && tag.length > bestTag.length))) {
      bestIndex = idx;
      bestTag = tag;
    }
  }
  return bestIndex === -1 ? null : { index: bestIndex, tag: bestTag };
}

function computeTrailingTagPrefixLength(lowerText: string, tags: string[]): number {
  const maxTagLen = Math.max(...tags.map((t) => t.length));
  const checkStart = Math.max(0, lowerText.length - (maxTagLen - 1));
  for (let i = checkStart; i < lowerText.length; i++) {
    const suffix = lowerText.slice(i);
    if (tags.some((t) => t.startsWith(suffix))) {
      return lowerText.length - i;
    }
  }
  return 0;
}

