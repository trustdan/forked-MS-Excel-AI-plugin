import {
  ProviderError,
  type ChatRequestOptions,
  type ChatResponse,
  type Provider,
  type ToolCall,
} from './types.js';
import { extractThinking } from './embedded.js';
import {
  MAX_ACCUMULATED_TEXT_CHARS,
  ThinkStreamSplitter,
  detectReasoningRumination,
  readSseStream,
  trimRuminatingReasoning,
} from './sse.js';
import { assertLoopbackUrl, guardedFetch, type Transport } from '../networkGuard.js';

/**
 * OpenAI Chat Completions adapter. With a custom baseUrl this same adapter
 * speaks to Ollama, LM Studio, vLLM, llama.cpp server, LiteLLM and any other
 * OpenAI-compatible endpoint — that's what makes "any model" true.
 *
 * Newer OpenAI models renamed `max_tokens` to `max_completion_tokens` and
 * some reject `temperature` overrides entirely, while most compatible
 * servers still expect the classic parameters. Rather than hard-coding model
 * lists, the adapter retries a 400 that names an offending parameter with an
 * adjusted body and remembers the working shape for the rest of the session.
 */
export class OpenAIProvider implements Provider {
  readonly supportsTools = true;

  private maxTokensParam: 'max_tokens' | 'max_completion_tokens' = 'max_tokens';
  private sendTemperature = true;
  private sendReasoningEffort = true;
  private sendStream = true;

  constructor(
    readonly model: string,
    private readonly apiKey: string | undefined = 'lm-studio',
    private readonly baseUrl = 'http://127.0.0.1:1234/v1',
    readonly id: string = 'openai',
    private readonly transport: Transport = guardedFetch,
  ) {
    assertLoopbackUrl(this.baseUrl);
  }

  async chat(options: ChatRequestOptions): Promise<ChatResponse> {
    // At most one retry per adaptable parameter.
    for (let attempt = 0; attempt < 5; attempt++) {
      const wantStream = Boolean(options.onDelta) && this.sendStream;
      const res = await this.send(options, wantStream);
      if (res.ok) {
        const contentType = (res.headers.get('content-type') ?? '').toLowerCase();
        if (wantStream && contentType.includes('text/event-stream')) {
          return this.parseStream(res, options);
        }
        return this.parse(res, options);
      }

      const text = await res.text().catch(() => '');
      if (res.status === 400 && this.adaptParams(text, wantStream)) continue;
      if (wantStream && (res.status === 404 || res.status === 405 || res.status === 422 || res.status === 501)) {
        this.sendStream = false;
        continue;
      }

      throw new ProviderError(
        `${this.id} API error ${res.status}: ${truncate(text)}`,
        res.status,
        res.status === 429 || res.status >= 500,
      );
    }
    throw new ProviderError(`${this.id}: request failed after parameter adaptation`);
  }

  /** Returns true if the 400 named a parameter we can adjust. */
  private adaptParams(errorText: string, attemptedStream = false): boolean {
    const lower = errorText.toLowerCase();
    const unsupported = /unsupported|not supported|does not support|invalid|unknown|unrecognized/.test(lower);
    if (!unsupported) return false;
    if (this.maxTokensParam === 'max_tokens' && lower.includes('max_tokens')) {
      this.maxTokensParam = 'max_completion_tokens';
      return true;
    }
    if (this.sendTemperature && lower.includes('temperature')) {
      this.sendTemperature = false;
      return true;
    }
    if (this.sendReasoningEffort && lower.includes('reasoning_effort')) {
      this.sendReasoningEffort = false;
      return true;
    }
    if (attemptedStream && this.sendStream && lower.includes('stream')) {
      this.sendStream = false;
      return true;
    }
    return false;
  }

  private async send(options: ChatRequestOptions, stream: boolean): Promise<Response> {
    const messages: unknown[] = options.messages.map((m) => {
      switch (m.role) {
        case 'system':
          return { role: 'system', content: m.content };
        case 'user':
          return { role: 'user', content: m.content };
        case 'assistant': {
          const out: Record<string, unknown> = { role: 'assistant', content: m.content || null };
          if (m.toolCalls && m.toolCalls.length > 0) {
            out.tool_calls = m.toolCalls.map((tc) => ({
              id: tc.id,
              type: 'function',
              function: { name: tc.name, arguments: tc.argsJson },
            }));
          }
          return out;
        }
        case 'tool':
          return { role: 'tool', tool_call_id: m.toolCallId, content: m.content };
      }
    });

    const body: Record<string, unknown> = {
      model: this.model,
      messages,
      [this.maxTokensParam]: options.maxTokens ?? 8192,
    };
    if (stream) {
      body.stream = true;
    }
    if (this.sendTemperature && options.temperature !== undefined) {
      body.temperature = options.temperature;
    }
    if (this.sendReasoningEffort && options.reasoningEffort !== undefined) {
      body.reasoning_effort = options.reasoningEffort;
    }
    if (options.tools && options.tools.length > 0) {
      body.tools = options.tools.map((t) => ({
        type: 'function',
        function: { name: t.name, description: t.description, parameters: t.inputSchema },
      }));
    }

    const headers: Record<string, string> = { 'content-type': 'application/json' };
    const key = this.apiKey || 'lm-studio';
    headers.authorization = `Bearer ${key}`;

    const endpoint = `${this.baseUrl.replace(/\/$/, '')}/chat/completions`;

    return this.transport(endpoint, {
      method: 'POST',
      headers,
      body: JSON.stringify(body),
      signal: options.abortSignal ?? null,
    });
  }

  private async parse(res: Response, options: ChatRequestOptions): Promise<ChatResponse> {
    const data = (await res.json()) as {
      choices?: {
        message?: {
          content?: string | null;
          reasoning_content?: string | null;
          reasoning?: string | null;
          tool_calls?: { id: string; function: { name: string; arguments: string } }[];
        };
        finish_reason?: string;
      }[];
      usage?: {
        prompt_tokens?: number;
        completion_tokens?: number;
        completion_tokens_details?: { reasoning_tokens?: number };
      };
    };

    const choice = data.choices?.[0];
    if (!choice?.message) throw new ProviderError(`${this.id}: empty response`);

    const toolCalls: ToolCall[] = (choice.message.tool_calls ?? []).map((tc) => ({
      id: tc.id,
      name: tc.function.name,
      argsJson: tc.function.arguments,
    }));

    const rawContent = choice.message.content ?? '';
    let text = rawContent;
    let inlineReasoning = '';
    if (/<(?:think|thinking)>/i.test(rawContent)) {
      const extracted = extractThinking(rawContent);
      text = extracted.cleanText;
      inlineReasoning = extracted.reasoning;
    }

    const explicitReasoning = (choice.message.reasoning_content ?? choice.message.reasoning ?? '').trim();
    const rawReasoningText = [explicitReasoning, inlineReasoning].filter(Boolean).join('\n\n');
    const isRuminating =
      !text.trim() && toolCalls.length === 0 && detectReasoningRumination(rawReasoningText);
    const reasoningText = isRuminating
      ? trimRuminatingReasoning(rawReasoningText)
      : rawReasoningText;
    const reasoningTokens = data.usage?.completion_tokens_details?.reasoning_tokens ?? 0;
    const outputTokens = data.usage?.completion_tokens ?? reasoningTokens;

    if (options.onDelta) {
      if (reasoningText) options.onDelta({ reasoning: reasoningText });
      if (text && !/<tool_call>/i.test(text)) options.onDelta({ text });
    }

    return this.buildChatResponse({
      text,
      toolCalls,
      reasoningText,
      finishReason: isRuminating ? 'length' : choice.finish_reason,
      promptTokens: data.usage?.prompt_tokens ?? 0,
      outputTokens,
      reasoningTokens: isRuminating ? Math.max(reasoningTokens, 1) : reasoningTokens,
    });
  }

  private async parseStream(res: Response, options: ChatRequestOptions): Promise<ChatResponse> {
    let rawContent = '';
    let explicitReasoning = '';
    let finishReason: string | undefined;
    let promptTokens = 0;
    let completionTokens: number | undefined;
    let reasoningTokens = 0;
    let cutOffRumination = false;

    const toolCallMap = new Map<number, { id: string; name: string; arguments: string }>();
    const splitter = new ThinkStreamSplitter(options.onDelta);
    const streamAbort = new AbortController();
    const onCallerAbort = (): void => streamAbort.abort();
    options.abortSignal?.addEventListener('abort', onCallerAbort, { once: true });

    try {
      await readSseStream(
        res,
        (ev) => {
          let chunk: {
            choices?: {
              delta?: {
                content?: string | null;
                reasoning_content?: string | null;
                reasoning?: string | null;
                tool_calls?: {
                  index?: number;
                  id?: string;
                  function?: { name?: string; arguments?: string };
                }[];
              };
              finish_reason?: string | null;
            }[];
            usage?: {
              prompt_tokens?: number;
              completion_tokens?: number;
              completion_tokens_details?: { reasoning_tokens?: number };
            };
            error?: { message?: string };
          };
          try {
            chunk = JSON.parse(ev.data);
          } catch {
            return;
          }

          if (chunk.error?.message) {
            throw new ProviderError(`${this.id} stream error: ${chunk.error.message}`);
          }

          if (chunk.usage) {
            if (chunk.usage.prompt_tokens !== undefined) promptTokens = chunk.usage.prompt_tokens;
            if (chunk.usage.completion_tokens !== undefined) {
              completionTokens = chunk.usage.completion_tokens;
            }
            if (chunk.usage.completion_tokens_details?.reasoning_tokens !== undefined) {
              reasoningTokens = chunk.usage.completion_tokens_details.reasoning_tokens;
            }
          }

          const choice = chunk.choices?.[0];
          if (!choice) return;

          if (choice.finish_reason) {
            finishReason = choice.finish_reason;
          }

          const delta = choice.delta;
          if (!delta) return;

          const reasoningPiece = delta.reasoning_content ?? delta.reasoning;
          if (reasoningPiece) {
            explicitReasoning = appendBounded(explicitReasoning, reasoningPiece);
            splitter.pushReasoning(reasoningPiece);
          }

          if (delta.content) {
            rawContent = appendBounded(rawContent, delta.content);
            splitter.pushContent(delta.content);
          }

          if (delta.tool_calls) {
            for (let i = 0; i < delta.tool_calls.length; i++) {
              const tc = delta.tool_calls[i]!;
              const idx = tc.index ?? i;
              let entry = toolCallMap.get(idx);
              if (!entry) {
                entry = { id: tc.id ?? `call_${idx}`, name: '', arguments: '' };
                toolCallMap.set(idx, entry);
              } else if (tc.id) {
                entry.id = tc.id;
              }
              if (tc.function?.name) {
                entry.name += tc.function.name;
                options.onDelta?.({ toolCallName: entry.name });
              }
              if (tc.function?.arguments) {
                entry.arguments = appendBounded(entry.arguments, tc.function.arguments);
              }
            }
          }

          const activeReasoning =
            explicitReasoning.length > 0
              ? explicitReasoning
              : splitter.getAccumulatedReasoning();
          if (
            toolCallMap.size === 0 &&
            !cutOffRumination &&
            detectReasoningRumination(activeReasoning)
          ) {
            cutOffRumination = true;
            finishReason = 'length';
            reasoningTokens = Math.max(reasoningTokens, 1);
            streamAbort.abort();
          }
        },
        { abortSignal: streamAbort.signal },
      );
    } catch (err) {
      // If a local server emits a runaway SSE line, buffer overflow, or circular reasoning rumination,
      // salvage the reasoning/content streamed so far and mark finishReason as 'length' so the agent
      // loop can rescue or retry immediately instead of stalling or crashing.
      if (cutOffRumination && !options.abortSignal?.aborted) {
        finishReason = 'length';
        reasoningTokens = Math.max(reasoningTokens, 1);
      } else if (err instanceof ProviderError && err.status === 413) {
        finishReason = 'length';
        reasoningTokens = Math.max(reasoningTokens, 1);
        for (const [idx, entry] of toolCallMap.entries()) {
          try {
            JSON.parse(entry.arguments);
          } catch {
            toolCallMap.delete(idx);
          }
        }
      } else {
        throw err;
      }
    } finally {
      options.abortSignal?.removeEventListener('abort', onCallerAbort);
    }

    splitter.flush();

    const toolCalls: ToolCall[] = Array.from(toolCallMap.entries())
      .sort(([a], [b]) => a - b)
      .map(([, tc]) => ({
        id: tc.id,
        name: tc.name,
        argsJson: tc.arguments,
      }))
      .filter((tc) => tc.name.length > 0);

    let text = rawContent;
    let inlineReasoning = '';
    if (/<(?:think|thinking)>/i.test(rawContent)) {
      const extracted = extractThinking(rawContent);
      text = extracted.cleanText;
      inlineReasoning = extracted.reasoning;
    }

    const rawReasoningText = [explicitReasoning.trim(), inlineReasoning.trim()]
      .filter(Boolean)
      .join('\n\n');
    const reasoningText = cutOffRumination
      ? trimRuminatingReasoning(rawReasoningText)
      : rawReasoningText;
    const estimatedTokens = Math.ceil((rawContent.length + explicitReasoning.length) / 4);
    const outputTokens = completionTokens ?? (reasoningTokens > 0 ? reasoningTokens : estimatedTokens);

    return this.buildChatResponse({
      text,
      toolCalls,
      reasoningText,
      finishReason: finishReason ?? (toolCalls.length > 0 ? 'tool_calls' : 'stop'),
      promptTokens,
      outputTokens,
      reasoningTokens,
    });
  }

  private buildChatResponse(params: {
    text: string;
    toolCalls: ToolCall[];
    reasoningText: string;
    finishReason?: string;
    promptTokens: number;
    outputTokens: number;
    reasoningTokens: number;
  }): ChatResponse {
    const {
      text,
      toolCalls,
      reasoningText,
      finishReason,
      promptTokens,
      outputTokens,
      reasoningTokens,
    } = params;

    // Local reasoning servers (LM Studio / llama.cpp) often return finish_reason: "stop"
    // with empty content and empty tool_calls when internal reasoning_content hits the
    // token or context ceiling mid-thought. Classify that as max_tokens so the agent
    // loop can recover or report token exhaustion accurately.
    const exhaustedInReasoning =
      !text.trim() &&
      toolCalls.length === 0 &&
      (reasoningText.trim().length > 0 || reasoningTokens > 0);

    const out: ChatResponse = {
      text,
      toolCalls,
      usage: {
        inputTokens: promptTokens,
        outputTokens,
      },
      stopReason:
        finishReason === 'tool_calls'
          ? 'tool_use'
          : finishReason === 'length' || exhaustedInReasoning
            ? 'max_tokens'
            : finishReason === 'stop'
              ? 'end'
              : 'other',
    };
    if (reasoningText.trim().length > 0) {
      out.reasoning = reasoningText;
    }
    return out;
  }
}

function appendBounded(current: string, addition: string): string {
  if (current.length + addition.length > MAX_ACCUMULATED_TEXT_CHARS) {
    throw new ProviderError(
      `Streamed response exceeded maximum buffer limit (${MAX_ACCUMULATED_TEXT_CHARS} chars)`,
      413,
      false,
    );
  }
  return current + addition;
}

function truncate(s: string, n = 400): string {
  return s.length > n ? `${s.slice(0, n)}…` : s;
}
