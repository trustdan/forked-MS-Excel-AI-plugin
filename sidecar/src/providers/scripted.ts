import type {
  ChatRequestOptions,
  ChatResponse,
  Provider,
  ToolCall,
} from './types.js';

export type ScriptedTurn =
  | {
      text: string;
      stopReason?: ChatResponse['stopReason'];
      reasoning?: string;
      reasoningChunks?: string[];
      textChunks?: string[];
    }
  | {
      toolCalls: { name: string; args: unknown }[];
      text?: string;
      reasoning?: string;
      reasoningChunks?: string[];
      textChunks?: string[];
    };

/**
 * Deterministic provider for tests and harness development: plays back a
 * fixed script of turns. Also useful for simulating misbehaving models
 * (malformed JSON, unknown tools) to exercise the repair-retry path.
 */
export class ScriptedProvider implements Provider {
  readonly id = 'scripted';
  readonly model = 'scripted-v1';
  readonly supportsTools = true;
  private turn = 0;
  readonly requests: ChatRequestOptions[] = [];

  constructor(private readonly script: ScriptedTurn[]) {}

  async chat(options: ChatRequestOptions): Promise<ChatResponse> {
    // Snapshot the message list: the agent loop mutates it in place between
    // turns, and recorded requests must reflect what this call actually saw.
    this.requests.push({ ...options, messages: [...options.messages] });
    const current = this.script[this.turn];
    if (!current) {
      return {
        text: 'Done.',
        toolCalls: [],
        usage: { inputTokens: 1, outputTokens: 1 },
        stopReason: 'end',
      };
    }
    this.turn++;
    if (options.onDelta) {
      for (const r of current.reasoningChunks ?? (current.reasoning ? [current.reasoning] : [])) {
        options.onDelta({ reasoning: r });
      }
      for (const t of current.textChunks ?? (current.text ? [current.text] : [])) {
        if (!/<(?:think|thinking|tool_call)>/i.test(t)) {
          options.onDelta({ text: t });
        }
      }
    }
    if ('toolCalls' in current) {
      const toolCalls: ToolCall[] = current.toolCalls.map((tc, i) => ({
        id: `call_${this.turn}_${i}`,
        name: tc.name,
        argsJson: typeof tc.args === 'string' ? tc.args : JSON.stringify(tc.args),
      }));
      return {
        text: current.text ?? '',
        toolCalls,
        usage: { inputTokens: 1, outputTokens: 1 },
        stopReason: 'tool_use',
        ...(current.reasoning !== undefined ? { reasoning: current.reasoning } : {}),
      };
    }
    const out: ChatResponse = {
      text: current.text,
      toolCalls: [],
      usage: { inputTokens: 1, outputTokens: 1 },
      stopReason: current.stopReason ?? 'end',
    };
    if (current.reasoning !== undefined) {
      out.reasoning = current.reasoning;
    }
    return out;
  }
}
