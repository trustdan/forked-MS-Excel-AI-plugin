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
  readSseStream,
} from './sse.js';
import { assertLoopbackUrl, guardedFetch, type Transport } from '../networkGuard.js';

/** Anthropic paths are rooted at `/v1`; accept a base URL given with or without it. */
export function anthropicUrl(baseUrl: string, path: 'messages' | 'models'): string {
  const base = baseUrl.replace(/\/+$/, '');
  return `${base.endsWith('/v1') ? base : `${base}/v1`}/${path}`;
}

interface AnthropicContentBlock {
  type: string;
  text?: string;
  thinking?: string;
  id?: string;
  name?: string;
  input?: unknown;
}

/** Anthropic Messages API adapter (plain fetch, no SDK). */
export class AnthropicProvider implements Provider {
  readonly id = 'anthropic';
  readonly supportsTools = true;
  private sendStream = true;

  constructor(
    readonly model: string,
    private readonly apiKey: string = 'lm-studio',
    private readonly baseUrl = 'http://127.0.0.1:1234/v1',
    private readonly transport: Transport = guardedFetch,
  ) {
    assertLoopbackUrl(this.baseUrl);
  }

  async chat(options: ChatRequestOptions): Promise<ChatResponse> {
    for (let attempt = 0; attempt < 2; attempt++) {
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
      if (
        wantStream &&
        (res.status === 400 || res.status === 404 || res.status === 405 || res.status === 422 || res.status === 501) &&
        /stream|unsupported|not supported|invalid/i.test(text)
      ) {
        this.sendStream = false;
        continue;
      }

      throw new ProviderError(
        `Anthropic API error ${res.status}: ${truncate(text)}`,
        res.status,
        res.status === 429 || res.status >= 500,
      );
    }
    throw new ProviderError('Anthropic: request failed after streaming fallback');
  }

  private async send(options: ChatRequestOptions, stream: boolean): Promise<Response> {
    const system = options.messages.find((m) => m.role === 'system');
    const messages: unknown[] = [];
    for (const m of options.messages) {
      if (m.role === 'system') continue;
      if (m.role === 'user') {
        messages.push({ role: 'user', content: m.content });
      } else if (m.role === 'assistant') {
        const content: AnthropicContentBlock[] = [];
        if (m.content) content.push({ type: 'text', text: m.content });
        for (const tc of m.toolCalls ?? []) {
          content.push({
            type: 'tool_use',
            id: tc.id,
            name: tc.name,
            input: safeParse(tc.argsJson),
          });
        }
        messages.push({ role: 'assistant', content });
      } else {
        messages.push({
          role: 'user',
          content: [{ type: 'tool_result', tool_use_id: m.toolCallId, content: m.content }],
        });
      }
    }

    const body: Record<string, unknown> = {
      model: this.model,
      max_tokens: options.maxTokens ?? 4096,
      messages,
    };
    if (stream) body.stream = true;
    if (system) body.system = system.content;
    if (options.temperature !== undefined) body.temperature = options.temperature;
    if (options.tools && options.tools.length > 0) {
      body.tools = options.tools.map((t) => ({
        name: t.name,
        description: t.description,
        input_schema: t.inputSchema,
      }));
    }

    return this.transport(anthropicUrl(this.baseUrl, 'messages'), {
      method: 'POST',
      headers: {
        'content-type': 'application/json',
        'x-api-key': this.apiKey || 'lm-studio',
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify(body),
      signal: options.abortSignal ?? null,
    });
  }

  private async parse(res: Response, options: ChatRequestOptions): Promise<ChatResponse> {
    const data = (await res.json()) as {
      content: AnthropicContentBlock[];
      stop_reason: string;
      usage?: { input_tokens?: number; output_tokens?: number };
    };

    let text = '';
    const reasoningParts: string[] = [];
    const toolCalls: ToolCall[] = [];
    for (const block of data.content ?? []) {
      if (block.type === 'thinking') {
        const thought = (block.thinking ?? block.text ?? '').trim();
        if (thought) reasoningParts.push(thought);
      } else if (block.type === 'text' && block.text) {
        text += block.text;
      } else if (block.type === 'tool_use' && block.id && block.name) {
        toolCalls.push({ id: block.id, name: block.name, argsJson: JSON.stringify(block.input ?? {}) });
      }
    }

    if (/<(?:think|thinking)>/i.test(text)) {
      const extracted = extractThinking(text);
      text = extracted.cleanText;
      if (extracted.reasoning) reasoningParts.push(extracted.reasoning);
    }

    const reasoningText = reasoningParts.join('\n\n').trim();

    if (options.onDelta) {
      if (reasoningText) options.onDelta({ reasoning: reasoningText });
      if (text && !/<tool_call>/i.test(text)) options.onDelta({ text });
    }

    const out: ChatResponse = {
      text,
      toolCalls,
      usage: {
        inputTokens: data.usage?.input_tokens ?? 0,
        outputTokens: data.usage?.output_tokens ?? 0,
      },
      stopReason:
        data.stop_reason === 'tool_use'
          ? 'tool_use'
          : data.stop_reason === 'max_tokens'
            ? 'max_tokens'
            : data.stop_reason === 'end_turn'
              ? 'end'
              : 'other',
    };
    if (reasoningText.length > 0) {
      out.reasoning = reasoningText;
    }
    return out;
  }

  private async parseStream(res: Response, options: ChatRequestOptions): Promise<ChatResponse> {
    let rawText = '';
    let explicitReasoning = '';
    let stopReason = 'end_turn';
    let inputTokens = 0;
    let outputTokens = 0;

    const toolBlocks = new Map<number, { id: string; name: string; inputJson: string }>();
    const splitter = new ThinkStreamSplitter(options.onDelta);

    await readSseStream(
      res,
      (ev) => {
        let payload: {
          type?: string;
          index?: number;
          message?: {
            stop_reason?: string | null;
            usage?: { input_tokens?: number; output_tokens?: number };
          };
          content_block?: {
            type?: string;
            id?: string;
            name?: string;
            text?: string;
            thinking?: string;
            input?: unknown;
          };
          delta?: {
            type?: string;
            text?: string;
            thinking?: string;
            partial_json?: string;
            stop_reason?: string | null;
          };
          usage?: { input_tokens?: number; output_tokens?: number };
          error?: { message?: string };
        };
        try {
          payload = JSON.parse(ev.data);
        } catch {
          return;
        }

        const eventType = payload.type ?? ev.event;
        if (eventType === 'error' || payload.error?.message) {
          throw new ProviderError(
            `Anthropic stream error: ${payload.error?.message ?? 'unknown error'}`,
          );
        }

        if (eventType === 'message_start') {
          if (payload.message?.usage?.input_tokens !== undefined) {
            inputTokens = payload.message.usage.input_tokens;
          }
          if (payload.message?.usage?.output_tokens !== undefined) {
            outputTokens = payload.message.usage.output_tokens;
          }
          return;
        }

        if (eventType === 'content_block_start' && payload.content_block) {
          const idx = payload.index ?? 0;
          const cb = payload.content_block;
          if (cb.type === 'tool_use' && cb.id && cb.name) {
            const initialJson =
              cb.input && typeof cb.input === 'object' && Object.keys(cb.input as object).length > 0
                ? JSON.stringify(cb.input)
                : '';
            toolBlocks.set(idx, { id: cb.id, name: cb.name, inputJson: initialJson });
          } else if (cb.type === 'thinking' && cb.thinking) {
            explicitReasoning = appendBounded(explicitReasoning, cb.thinking);
            splitter.pushReasoning(cb.thinking);
          } else if (cb.type === 'text' && cb.text) {
            rawText = appendBounded(rawText, cb.text);
            splitter.pushContent(cb.text);
          }
          return;
        }

        if (eventType === 'content_block_delta' && payload.delta) {
          const idx = payload.index ?? 0;
          const d = payload.delta;
          if (d.type === 'thinking_delta' && d.thinking) {
            explicitReasoning = appendBounded(explicitReasoning, d.thinking);
            splitter.pushReasoning(d.thinking);
          } else if (d.type === 'text_delta' && d.text) {
            rawText = appendBounded(rawText, d.text);
            splitter.pushContent(d.text);
          } else if (d.type === 'input_json_delta' && d.partial_json !== undefined) {
            const entry = toolBlocks.get(idx);
            if (entry) {
              entry.inputJson = appendBounded(entry.inputJson, d.partial_json);
            }
          }
          return;
        }

        if (eventType === 'message_delta') {
          if (payload.delta?.stop_reason) {
            stopReason = payload.delta.stop_reason;
          }
          if (payload.usage?.input_tokens !== undefined) {
            inputTokens = payload.usage.input_tokens;
          }
          if (payload.usage?.output_tokens !== undefined) {
            outputTokens = payload.usage.output_tokens;
          }
        }
      },
      { abortSignal: options.abortSignal },
    );

    splitter.flush();

    const toolCalls: ToolCall[] = Array.from(toolBlocks.entries())
      .sort(([a], [b]) => a - b)
      .map(([, tb]) => ({
        id: tb.id,
        name: tb.name,
        argsJson: tb.inputJson.trim() ? tb.inputJson : '{}',
      }));

    let text = rawText;
    let inlineReasoning = '';
    if (/<(?:think|thinking)>/i.test(text)) {
      const extracted = extractThinking(text);
      text = extracted.cleanText;
      inlineReasoning = extracted.reasoning;
    }

    const reasoningText = [explicitReasoning.trim(), inlineReasoning.trim()]
      .filter(Boolean)
      .join('\n\n');

    if (outputTokens === 0 && (rawText.length > 0 || explicitReasoning.length > 0)) {
      outputTokens = Math.ceil((rawText.length + explicitReasoning.length) / 4);
    }

    const out: ChatResponse = {
      text,
      toolCalls,
      usage: {
        inputTokens,
        outputTokens,
      },
      stopReason:
        stopReason === 'tool_use' || toolCalls.length > 0
          ? 'tool_use'
          : stopReason === 'max_tokens'
            ? 'max_tokens'
            : stopReason === 'end_turn'
              ? 'end'
              : 'other',
    };
    if (reasoningText.length > 0) {
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

function safeParse(json: string): unknown {
  try {
    return JSON.parse(json);
  } catch {
    return {};
  }
}

function truncate(s: string, n = 400): string {
  return s.length > n ? `${s.slice(0, n)}…` : s;
}
