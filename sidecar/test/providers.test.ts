import { afterEach, describe, expect, it, vi } from 'vitest';
import { AnthropicProvider } from '../src/providers/anthropic.js';
import { OpenAIProvider } from '../src/providers/openai.js';
import {
  extractThinking,
  readSseStream,
  ThinkStreamSplitter,
  withRetry,
} from '../src/providers/index.js';
import { ProviderError, type ChatDelta } from '../src/providers/types.js';

afterEach(() => {
  vi.unstubAllGlobals();
});

function stubFetch(handler: (url: string, init: RequestInit) => { status: number; body: unknown }) {
  const calls: { url: string; body: unknown }[] = [];
  vi.stubGlobal('fetch', async (url: string | URL, init?: RequestInit) => {
    const parsedBody = init?.body ? JSON.parse(String(init.body)) : undefined;
    calls.push({ url: String(url), body: parsedBody });
    const res = handler(String(url), init ?? {});
    return new Response(JSON.stringify(res.body), {
      status: res.status,
      headers: { 'content-type': 'application/json' },
    });
  });
  return calls;
}

describe('Anthropic adapter', () => {
  it('maps tool definitions, tool calls and usage', async () => {
    const calls = stubFetch(() => ({
      status: 200,
      body: {
        content: [
          { type: 'text', text: 'Reading the range.' },
          { type: 'tool_use', id: 'tu_1', name: 'read_range', input: { range: 'Sheet1!A1' } },
        ],
        stop_reason: 'tool_use',
        usage: { input_tokens: 11, output_tokens: 7 },
      },
    }));

    const p = new AnthropicProvider('claude-test', 'sk-test');
    const res = await p.chat({
      messages: [
        { role: 'system', content: 'sys' },
        { role: 'user', content: 'hello' },
      ],
      tools: [{ name: 'read_range', description: 'd', inputSchema: { type: 'object' } }],
    });

    expect(res.toolCalls).toEqual([
      { id: 'tu_1', name: 'read_range', argsJson: '{"range":"Sheet1!A1"}' },
    ]);
    expect(res.stopReason).toBe('tool_use');
    expect(res.usage).toEqual({ inputTokens: 11, outputTokens: 7 });

    const sent = calls[0]!.body as Record<string, unknown>;
    expect(sent.system).toBe('sys');
    expect((sent.tools as unknown[])[0]).toMatchObject({ name: 'read_range' });
    expect(calls[0]!.url).toContain('/v1/messages');
  });

  it('throws a retryable ProviderError on 429/5xx', async () => {
    stubFetch(() => ({ status: 429, body: { error: 'rate limited' } }));
    const p = new AnthropicProvider('claude-test', 'sk-test');
    await expect(p.chat({ messages: [{ role: 'user', content: 'x' }] })).rejects.toMatchObject({
      retryable: true,
      status: 429,
    });
  });
});

describe('OpenAI / OpenAI-compatible adapter', () => {
  it('maps tool calls and respects custom base URLs (Ollama etc.)', async () => {
    const calls = stubFetch(() => ({
      status: 200,
      body: {
        choices: [
          {
            message: {
              content: null,
              tool_calls: [
                { id: 'c1', function: { name: 'read_range', arguments: '{"range":"Sheet1!A1"}' } },
              ],
            },
            finish_reason: 'tool_calls',
          },
        ],
        usage: { prompt_tokens: 5, completion_tokens: 3 },
      },
    }));

    const p = new OpenAIProvider('llama3', undefined, 'http://localhost:11434/v1', 'openai-compatible');
    const res = await p.chat({ messages: [{ role: 'user', content: 'x' }] });

    expect(calls[0]!.url).toBe('http://localhost:11434/v1/chat/completions');
    expect(res.toolCalls[0]).toMatchObject({ name: 'read_range' });
    expect(res.stopReason).toBe('tool_use');
  });

  it('round-trips assistant tool calls and tool results in history', async () => {
    const calls = stubFetch(() => ({
      status: 200,
      body: { choices: [{ message: { content: 'ok' }, finish_reason: 'stop' }] },
    }));
    const p = new OpenAIProvider('gpt-test', 'key');
    await p.chat({
      messages: [
        { role: 'user', content: 'q' },
        {
          role: 'assistant',
          content: '',
          toolCalls: [{ id: 'c1', name: 't', argsJson: '{}' }],
        },
        { role: 'tool', toolCallId: 'c1', content: '{"ok":true}' },
      ],
    });
    const sent = calls[0]!.body as { messages: Record<string, unknown>[] };
    expect(sent.messages[1]).toMatchObject({ role: 'assistant' });
    expect(sent.messages[2]).toMatchObject({ role: 'tool', tool_call_id: 'c1' });
  });

  it('classifies empty content with reasoning_content / reasoning_tokens as max_tokens', async () => {
    stubFetch(() => ({
      status: 200,
      body: {
        choices: [
          {
            message: {
              role: 'assistant',
              content: '',
              reasoning_content: 'Thinking through columns A1:E6... - E6',
              tool_calls: [],
            },
            finish_reason: 'stop',
          },
        ],
        usage: {
          completion_tokens_details: { reasoning_tokens: 5060 },
        },
      },
    }));

    const p = new OpenAIProvider('recursivesi/flint', 'lm-studio', 'http://127.0.0.1:1234/v1', 'openai-compatible');
    const res = await p.chat({ messages: [{ role: 'user', content: 'please continue' }] });
    expect(res.text).toBe('');
    expect(res.toolCalls).toEqual([]);
    expect(res.stopReason).toBe('max_tokens');
    expect(res.usage.outputTokens).toBe(5060);
    expect(res.reasoning).toBe('Thinking through columns A1:E6... - E6');
  });
});

describe('OpenAI parameter adaptation (newer models)', () => {
  it('switches max_tokens -> max_completion_tokens on the documented 400 and remembers it', async () => {
    const calls = stubFetch((_url, init) => {
      const body = JSON.parse(String(init.body)) as Record<string, unknown>;
      if ('max_tokens' in body) {
        return {
          status: 400,
          body: {
            error: {
              message:
                "Unsupported parameter: 'max_tokens' is not supported with this model. Use 'max_completion_tokens' instead.",
            },
          },
        };
      }
      return {
        status: 200,
        body: { choices: [{ message: { content: 'hi' }, finish_reason: 'stop' }] },
      };
    });

    const p = new OpenAIProvider('gpt-5.2', 'key');
    const res = await p.chat({ messages: [{ role: 'user', content: 'x' }], maxTokens: 100 });
    expect(res.text).toBe('hi');
    expect(calls).toHaveLength(2);
    expect(calls[1]!.body).toMatchObject({ max_completion_tokens: 100 });

    // Second request uses the learned parameter immediately — no extra 400.
    await p.chat({ messages: [{ role: 'user', content: 'y' }], maxTokens: 50 });
    expect(calls).toHaveLength(3);
    expect(calls[2]!.body).toMatchObject({ max_completion_tokens: 50 });
  });

  it('drops temperature when the model rejects it', async () => {
    const calls = stubFetch((_url, init) => {
      const body = JSON.parse(String(init.body)) as Record<string, unknown>;
      if ('temperature' in body) {
        return {
          status: 400,
          body: {
            error: { message: "Unsupported value: 'temperature' does not support 0 with this model." },
          },
        };
      }
      return {
        status: 200,
        body: { choices: [{ message: { content: 'ok' }, finish_reason: 'stop' }] },
      };
    });

    const p = new OpenAIProvider('o-reasoner', 'key');
    const res = await p.chat({ messages: [{ role: 'user', content: 'x' }], temperature: 0 });
    expect(res.text).toBe('ok');
    expect(calls).toHaveLength(2);
    expect('temperature' in (calls[1]!.body as Record<string, unknown>)).toBe(false);
  });

  it('sends reasoning_effort and drops it when the endpoint rejects it', async () => {
    const calls = stubFetch((_url, init) => {
      const body = JSON.parse(String(init.body)) as Record<string, unknown>;
      if ('reasoning_effort' in body) {
        return {
          status: 400,
          body: {
            error: { message: "Unsupported parameter: 'reasoning_effort' is not supported." },
          },
        };
      }
      return {
        status: 200,
        body: { choices: [{ message: { content: 'ok' }, finish_reason: 'stop' }] },
      };
    });

    const p = new OpenAIProvider('legacy-local', 'key');
    const res = await p.chat({ messages: [{ role: 'user', content: 'x' }], reasoningEffort: 'low' });
    expect(res.text).toBe('ok');
    expect(calls).toHaveLength(2);
    expect((calls[0]!.body as Record<string, unknown>).reasoning_effort).toBe('low');
    expect('reasoning_effort' in (calls[1]!.body as Record<string, unknown>)).toBe(false);
  });

  it('does not loop on unrelated 400s', async () => {
    const calls = stubFetch(() => ({
      status: 400,
      body: { error: { message: 'Invalid request: messages must not be empty' } },
    }));
    const p = new OpenAIProvider('gpt-test', 'key');
    await expect(p.chat({ messages: [{ role: 'user', content: 'x' }] })).rejects.toThrow('400');
    expect(calls.length).toBeLessThanOrEqual(2);
  });
});

describe('withRetry', () => {
  it('retries retryable errors with backoff then succeeds', async () => {
    let attempts = 0;
    const result = await withRetry(
      async () => {
        attempts++;
        if (attempts < 3) throw new ProviderError('rate', 429, true);
        return 'ok';
      },
      3,
      1,
    );
    expect(result).toBe('ok');
    expect(attempts).toBe(3);
  });

  it('does not retry non-retryable errors', async () => {
    let attempts = 0;
    await expect(
      withRetry(
        async () => {
          attempts++;
          throw new ProviderError('bad key', 401, false);
        },
        3,
        1,
      ),
    ).rejects.toThrow('bad key');
    expect(attempts).toBe(1);
  });
});


describe('Reasoning & <think> block normalization', () => {
  it('extracts closed and unclosed <think>/<thinking> blocks cleanly', () => {
    const r1 = extractThinking('<think>Step 1: check sheet\nStep 2: sum</think>The total is 42.');
    expect(r1.reasoning).toBe('Step 1: check sheet\nStep 2: sum');
    expect(r1.cleanText).toBe('The total is 42.');

    // Unclosed <think> before <tool_call> preserves the tool call in cleanText
    const r2 = extractThinking(
      '<think>I should inspect A1:B5 first\n<tool_call>{"name":"read_range","arguments":{"range":"Sheet1!A1:B5"}}</tool_call>',
    );
    expect(r2.reasoning).toBe('I should inspect A1:B5 first');
    expect(r2.cleanText).toContain('<tool_call>');
  });

  it('strips inline <think> tags in OpenAIProvider and merges with reasoning_content', async () => {
    stubFetch(() => ({
      status: 200,
      body: {
        choices: [
          {
            message: {
              role: 'assistant',
              content: '<think>Evaluating Q1 revenue columns</think>Here is the Q1 breakdown.',
            },
            finish_reason: 'stop',
          },
        ],
        usage: { prompt_tokens: 10, completion_tokens: 20 },
      },
    }));

    const p = new OpenAIProvider('deepseek-r1', 'lm-studio', 'http://127.0.0.1:1234/v1', 'openai-compatible');
    const res = await p.chat({ messages: [{ role: 'user', content: 'summarize Q1' }] });
    expect(res.text).toBe('Here is the Q1 breakdown.');
    expect(res.reasoning).toBe('Evaluating Q1 revenue columns');
  });

  it('extracts thinking content blocks in AnthropicProvider', async () => {
    stubFetch(() => ({
      status: 200,
      body: {
        content: [
          { type: 'thinking', thinking: 'Checking formula dependencies in B2:B10.' },
          { type: 'text', text: 'All formulas are valid.' },
        ],
        stop_reason: 'end_turn',
        usage: { input_tokens: 12, output_tokens: 18 },
      },
    }));

    const p = new AnthropicProvider('claude-local', 'sk-test');
    const res = await p.chat({ messages: [{ role: 'user', content: 'verify formulas' }] });
    expect(res.text).toBe('All formulas are valid.');
    expect(res.reasoning).toBe('Checking formula dependencies in B2:B10.');
  });
});

function createSseResponse(chunks: string[], status = 200): Response {
  const encoder = new TextEncoder();
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      for (const chunk of chunks) {
        controller.enqueue(encoder.encode(chunk));
      }
      controller.close();
    },
  });
  return new Response(stream, {
    status,
    headers: { 'content-type': 'text/event-stream' },
  });
}

describe('SSE Streaming & ThinkStreamSplitter (Sprint 6.2)', () => {
  it('splits incremental <think>...</think> and <tool_call> tags across token boundaries cleanly', () => {
    const deltas: ChatDelta[] = [];
    const splitter = new ThinkStreamSplitter((d) => deltas.push(d));

    // Tag split across chunks: "<thi" + "nk>Step 1" + "</th" + "ink>\nHello " + "world<tool_" + "call>{\"name\":\"x\"}"
    splitter.pushContent('<thi');
    splitter.pushContent('nk>Step 1: ');
    splitter.pushContent('check cells</th');
    splitter.pushContent('ink>\nHello ');
    splitter.pushContent('world!<tool_');
    splitter.pushContent('call>{"name":"read_range"}');
    splitter.flush();

    const reasoningStream = deltas
      .map((d) => d.reasoning ?? '')
      .filter(Boolean)
      .join('');
    const textStream = deltas
      .map((d) => d.text ?? '')
      .filter(Boolean)
      .join('');

    expect(reasoningStream).toBe('Step 1: check cells');
    expect(textStream).toBe('Hello world!');
  });

  it('enforces maxStreamBytes in readSseStream and cancels the reader on overflow', async () => {
    const res = createSseResponse([
      'data: {"choices":[{"delta":{"content":"1234567890"}}]}\n\n',
      'data: {"choices":[{"delta":{"content":"1234567890"}}]}\n\n',
    ]);

    await expect(
      readSseStream(res, () => {}, { maxStreamBytes: 20 }),
    ).rejects.toThrow(/exceeded maximum allowed size/);
  });

  it('streams reasoning_content, inline <think>, text deltas, and chunked tool_calls in OpenAIProvider', async () => {
    const requests: Record<string, unknown>[] = [];
    vi.stubGlobal('fetch', async (_url: string | URL, init?: RequestInit) => {
      requests.push(JSON.parse(String(init?.body)));
      return createSseResponse([
        'data: {"choices":[{"delta":{"reasoning_content":"Inspecting "}}]}\n\n',
        'data: {"choices":[{"delta":{"reasoning_content":"Sheet1!A1:B2."}}]}\n\n',
        'data: {"choices":[{"delta":{"content":"Reading cells now."}}]}\n\n',
        'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call_1","function":{"name":"read_range","arguments":"{\\"range\\":"}}]}}]}\n\n',
        'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"arguments":"\\"Sheet1!A1:B2\\"}"}}]},"finish_reason":"tool_calls"}],"usage":{"prompt_tokens":14,"completion_tokens":22}}\n\n',
        'data: [DONE]\n\n',
      ]);
    });

    const p = new OpenAIProvider('qwen-local', 'lm-studio', 'http://127.0.0.1:1234/v1', 'openai-compatible');
    const deltas: ChatDelta[] = [];
    const res = await p.chat({
      messages: [{ role: 'user', content: 'read A1:B2' }],
      onDelta: (d) => deltas.push(d),
    });

    expect(requests[0]?.stream).toBe(true);
    expect(deltas.map((d) => d.reasoning).filter(Boolean).join('')).toBe('Inspecting Sheet1!A1:B2.');
    expect(deltas.map((d) => d.text).filter(Boolean).join('')).toBe('Reading cells now.');
    expect(res.reasoning).toBe('Inspecting Sheet1!A1:B2.');
    expect(res.text).toBe('Reading cells now.');
    expect(res.stopReason).toBe('tool_use');
    expect(res.toolCalls).toEqual([
      { id: 'call_1', name: 'read_range', argsJson: '{"range":"Sheet1!A1:B2"}' },
    ]);
    expect(res.usage).toEqual({ inputTokens: 14, outputTokens: 22 });
  });

  it('automatically falls back to non-streaming when a local server rejects stream: true', async () => {
    const requests: Record<string, unknown>[] = [];
    vi.stubGlobal('fetch', async (_url: string | URL, init?: RequestInit) => {
      const body = JSON.parse(String(init?.body)) as Record<string, unknown>;
      requests.push(body);
      if (body.stream) {
        return new Response(
          JSON.stringify({ error: { message: "Unsupported parameter: 'stream' is not supported." } }),
          { status: 400, headers: { 'content-type': 'application/json' } },
        );
      }
      return new Response(
        JSON.stringify({
          choices: [{ message: { content: 'Fallback response ok' }, finish_reason: 'stop' }],
          usage: { prompt_tokens: 5, completion_tokens: 4 },
        }),
        { status: 200, headers: { 'content-type': 'application/json' } },
      );
    });

    const p = new OpenAIProvider('no-stream-model', 'lm-studio');
    const deltas: ChatDelta[] = [];
    const res = await p.chat({
      messages: [{ role: 'user', content: 'hi' }],
      onDelta: (d) => deltas.push(d),
    });

    expect(requests).toHaveLength(2);
    expect(requests[0]?.stream).toBe(true);
    expect(requests[1]?.stream).toBeUndefined();
    expect(res.text).toBe('Fallback response ok');
    expect(deltas).toEqual([{ text: 'Fallback response ok' }]);
  });

  it('streams thinking_delta, text_delta, and input_json_delta in AnthropicProvider', async () => {
    vi.stubGlobal('fetch', async () =>
      createSseResponse([
        'event: message_start\ndata: {"type":"message_start","message":{"usage":{"input_tokens":15,"output_tokens":1}}}\n\n',
        'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"thinking","thinking":""}}\n\n',
        'event: content_block_delta\ndata: {"type":"content_block_delta","index":0,"delta":{"type":"thinking_delta","thinking":"Checking sheet "}}\n\n',
        'event: content_block_delta\ndata: {"type":"content_block_delta","index":0,"delta":{"type":"thinking_delta","thinking":"formulas."}}\n\n',
        'event: content_block_start\ndata: {"type":"content_block_start","index":1,"content_block":{"type":"text","text":""}}\n\n',
        'event: content_block_delta\ndata: {"type":"content_block_delta","index":1,"delta":{"type":"text_delta","text":"All good."}}\n\n',
        'event: content_block_start\ndata: {"type":"content_block_start","index":2,"content_block":{"type":"tool_use","id":"tu_9","name":"read_range","input":{}}}\n\n',
        'event: content_block_delta\ndata: {"type":"content_block_delta","index":2,"delta":{"type":"input_json_delta","partial_json":"{\\"range\\":\\"Sheet1!A1\\"}"}}\n\n',
        'event: message_delta\ndata: {"type":"message_delta","delta":{"stop_reason":"tool_use"},"usage":{"output_tokens":28}}\n\n',
      ]),
    );

    const p = new AnthropicProvider('claude-stream', 'sk-test');
    const deltas: ChatDelta[] = [];
    const res = await p.chat({
      messages: [{ role: 'user', content: 'check' }],
      onDelta: (d) => deltas.push(d),
    });

    expect(deltas.map((d) => d.reasoning).filter(Boolean).join('')).toBe('Checking sheet formulas.');
    expect(deltas.map((d) => d.text).filter(Boolean).join('')).toBe('All good.');
    expect(res.reasoning).toBe('Checking sheet formulas.');
    expect(res.text).toBe('All good.');
    expect(res.toolCalls).toEqual([
      { id: 'tu_9', name: 'read_range', argsJson: '{"range":"Sheet1!A1"}' },
    ]);
    expect(res.stopReason).toBe('tool_use');
    expect(res.usage).toEqual({ inputTokens: 15, outputTokens: 28 });
  });

  it('salvages streamed reasoning and returns stopReason max_tokens instead of crashing when an SSE stream hits a 413 buffer overflow', async () => {
    const giantLineWithoutNewline = 'data: {"choices":[{"delta":{"content":"' + 'x'.repeat(9 * 1024 * 1024);
    vi.stubGlobal('fetch', async () =>
      createSseResponse([
        'data: {"choices":[{"delta":{"reasoning_content":"Drafted balance sheet rows 1..33."}}]}\n\n',
        giantLineWithoutNewline,
      ]),
    );

    const p = new OpenAIProvider('qwen-local', 'lm-studio', 'http://127.0.0.1:1234/v1', 'openai-compatible');
    const deltas: ChatDelta[] = [];
    const res = await p.chat({
      messages: [{ role: 'user', content: 'draft balance sheet' }],
      onDelta: (d) => deltas.push(d),
    });

    expect(res.stopReason).toBe('max_tokens');
    expect(res.reasoning).toBe('Drafted balance sheet rows 1..33.');
  });
});
