/**
 * Provider-agnostic chat interface. Implementations translate to each
 * vendor's wire format with plain fetch — no vendor SDKs, so any
 * OpenAI-compatible endpoint (Ollama, LM Studio, vLLM, gateways) works with
 * the same adapter.
 */

export interface ToolSpec {
  name: string;
  description: string;
  /** JSON schema for the tool parameters. */
  inputSchema: Record<string, unknown>;
}

export type ChatMessage =
  | { role: 'system'; content: string }
  | { role: 'user'; content: string }
  | { role: 'assistant'; content: string; toolCalls?: ToolCall[] }
  | { role: 'tool'; toolCallId: string; content: string };

export interface ToolCall {
  id: string;
  name: string;
  /** Raw JSON arguments string as produced by the model. */
  argsJson: string;
}

export interface ChatUsage {
  inputTokens: number;
  outputTokens: number;
}

export interface ChatResponse {
  text: string;
  toolCalls: ToolCall[];
  usage: ChatUsage;
  stopReason: 'end' | 'tool_use' | 'max_tokens' | 'other';
  /** Optional internal chain-of-thought returned by reasoning models. */
  reasoning?: string;
}

export interface ChatDelta {
  reasoning?: string;
  text?: string;
  toolCallName?: string;
}

export interface ChatRequestOptions {
  messages: ChatMessage[];
  tools?: ToolSpec[];
  maxTokens?: number;
  temperature?: number;
  reasoningEffort?: 'low' | 'medium' | 'high';
  abortSignal?: AbortSignal;
  /** Optional callback invoked as streaming tokens/reasoning arrive via SSE or in-process callbacks. */
  onDelta?: (delta: ChatDelta) => void;
}

export interface Provider {
  readonly id: string;
  readonly model: string;
  /** Whether the underlying API supports native tool calling. */
  readonly supportsTools: boolean;
  chat(options: ChatRequestOptions): Promise<ChatResponse>;
}

export class ProviderError extends Error {
  constructor(
    message: string,
    public readonly status?: number,
    public readonly retryable = false,
  ) {
    super(message);
  }
}

export interface ProviderConfig {
  kind: 'anthropic' | 'openai' | 'openai-compatible' | 'embedded';
  model: string;
  apiKey?: string;
  /** Required for openai-compatible; defaults applied for the rest. */
  baseUrl?: string;
  /** Optional explicit path to .gguf file for embedded engine. */
  modelPath?: string;
  /** Context window size in tokens for embedded engine (default: 4096). */
  contextSize?: number;
  /** Number of CPU threads for embedded execution. */
  threads?: number;
  /** Number of layers to offload to GPU if supported. */
  gpuLayers?: number;
}
