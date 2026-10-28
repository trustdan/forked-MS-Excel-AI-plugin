import { existsSync, mkdirSync, readdirSync, statSync } from 'node:fs';
import { basename, isAbsolute, join, resolve } from 'node:path';
import { dataDir } from '../config.js';
import {
  ProviderError,
  type ChatMessage,
  type ChatRequestOptions,
  type ChatResponse,
  type Provider,
  type ProviderConfig,
  type ToolCall,
  type ToolSpec,
} from './types.js';
import { ThinkStreamSplitter } from './sse.js';

export function getModelsDir(): string {
  const custom = process.env.EXCELAI_MODELS_DIR;
  const dir = custom ? resolve(custom) : join(dataDir(), 'models');
  try {
    mkdirSync(dir, { recursive: true });
  } catch {
    // best-effort directory creation
  }
  return dir;
}

/**
 * Lists all .gguf model files present in the local models directory.
 */
export async function listLocalGgufModels(customDir?: string): Promise<string[]> {
  const dir = customDir ? resolve(customDir) : getModelsDir();
  if (!existsSync(dir)) return [];
  try {
    const entries = readdirSync(dir, { withFileTypes: true });
    return entries
      .filter((e) => e.isFile() && e.name.toLowerCase().endsWith('.gguf'))
      .map((e) => e.name)
      .sort((a, b) => a.localeCompare(b));
  } catch {
    return [];
  }
}

/**
 * Validates and safely resolves a model path, defending against path traversal.
 */
export function resolveModelPath(model: string, explicitPath?: string): string {
  const modelsDir = getModelsDir();
  let candidate: string;

  if (explicitPath) {
    candidate = resolve(explicitPath);
  } else if (isAbsolute(model)) {
    candidate = resolve(model);
  } else {
    // Defend against path traversal in relative names (e.g. '../../etc/passwd')
    const sanitized = basename(model);
    const withExt = sanitized.toLowerCase().endsWith('.gguf') ? sanitized : `${sanitized}.gguf`;
    candidate = join(modelsDir, withExt);
  }

  if (!candidate.toLowerCase().endsWith('.gguf')) {
    throw new ProviderError(
      `Invalid model file "${candidate}": embedded provider requires a .gguf model file.`,
      400,
    );
  }

  if (!existsSync(candidate)) {
    throw new ProviderError(
      `Embedded model file not found: "${candidate}". Place your .gguf model in "${modelsDir}" or specify an absolute path.`,
      404,
    );
  }

  try {
    const stat = statSync(candidate);
    if (!stat.isFile()) {
      throw new ProviderError(`Model path "${candidate}" is not a regular file.`, 400);
    }
  } catch (e) {
    if (e instanceof ProviderError) throw e;
    throw new ProviderError(`Failed to access model file "${candidate}": ${String(e)}`, 400);
  }

  return candidate;
}

/**
 * Backend interface for in-process execution.
 */
export interface LlamaEngineBackend {
  generate(
    prompt: string,
    options: {
      modelPath: string;
      contextSize?: number;
      threads?: number;
      gpuLayers?: number;
      temperature?: number;
      maxTokens?: number;
      abortSignal?: AbortSignal;
      onToken?: (chunk: string) => void;
    },
  ): Promise<{ text: string; inputTokens?: number; outputTokens?: number }>;
  dispose?(): Promise<void>;
}

/**
 * Dynamic node-llama-cpp backend. Loads bindings on demand to ensure
 * installation remains zero-script and builds never fail when native toolchains
 * are absent.
 */
export class NodeLlamaCppBackend implements LlamaEngineBackend {
  private static cachedLlama: unknown = null;
  private static cachedModels = new Map<string, unknown>();

  async generate(
    prompt: string,
    options: {
      modelPath: string;
      contextSize?: number;
      threads?: number;
      gpuLayers?: number;
      temperature?: number;
      maxTokens?: number;
      abortSignal?: AbortSignal;
      onToken?: (chunk: string) => void;
    },
  ): Promise<{ text: string; inputTokens?: number; outputTokens?: number }> {
    if (options.abortSignal?.aborted) {
      throw new Error('AbortError: Generation canceled by user');
    }

    let llamaModule: any;
    try {
      // Dynamic import to isolate optional native dependency
      // @ts-ignore
      llamaModule = await import('node-llama-cpp');
    } catch {
      throw new ProviderError(
        'node-llama-cpp native bindings are not installed or compiled in this environment. ' +
          'To use in-process GGUF execution, install node-llama-cpp or use the LM Studio / Ollama provider.',
        501,
      );
    }

    try {
      const getLlama = llamaModule.getLlama;
      if (!NodeLlamaCppBackend.cachedLlama) {
        NodeLlamaCppBackend.cachedLlama = await getLlama({
          gpu: options.gpuLayers && options.gpuLayers > 0 ? 'auto' : false,
        });
      }
      const llama = NodeLlamaCppBackend.cachedLlama as any;

      let model = NodeLlamaCppBackend.cachedModels.get(options.modelPath) as any;
      if (!model) {
        model = await llama.loadModel({
          modelPath: options.modelPath,
          gpuLayers: options.gpuLayers,
        });
        NodeLlamaCppBackend.cachedModels.set(options.modelPath, model);
      }

      const context = await model.createContext({
        contextSize: options.contextSize ?? 4096,
        threads: options.threads,
      });

      const session = new llamaModule.LlamaChatSession({
        contextSequence: context.getSequence(),
      });

      const response = await session.prompt(prompt, {
        temperature: options.temperature ?? 0.2,
        maxTokens: options.maxTokens ?? 2048,
        signal: options.abortSignal,
        onTextChunk: options.onToken
          ? (chunk: string) => {
              options.onToken?.(chunk);
            }
          : undefined,
      });

      return {
        text: response,
        inputTokens: Math.ceil(prompt.length / 4),
        outputTokens: Math.ceil(response.length / 4),
      };
    } catch (e: unknown) {
      if (options.abortSignal?.aborted || (e instanceof Error && e.name === 'AbortError')) {
        throw new Error('AbortError: Generation canceled by user');
      }
      throw new ProviderError(`Embedded GGUF execution error: ${e instanceof Error ? e.message : String(e)}`, 500);
    }
  }

  async dispose(): Promise<void> {
    NodeLlamaCppBackend.cachedModels.clear();
    NodeLlamaCppBackend.cachedLlama = null;
  }
}

/**
 * Deterministic mock backend for automated test suites and environments
 * without multi-gigabyte GGUF model files.
 */
export class MockLlamaBackend implements LlamaEngineBackend {
  constructor(
    private handler?: (
      prompt: string,
      signal?: AbortSignal,
      onToken?: (chunk: string) => void,
    ) => Promise<string> | string,
  ) {}

  async generate(
    prompt: string,
    options: {
      modelPath: string;
      temperature?: number;
      maxTokens?: number;
      abortSignal?: AbortSignal;
      onToken?: (chunk: string) => void;
    },
  ): Promise<{ text: string; inputTokens?: number; outputTokens?: number }> {
    if (options.abortSignal?.aborted) {
      throw new Error('AbortError: Generation canceled by user');
    }

    if (this.handler) {
      let emittedToken = false;
      const wrappedOnToken = (chunk: string): void => {
        emittedToken = true;
        options.onToken?.(chunk);
      };
      const result = await this.handler(prompt, options.abortSignal, wrappedOnToken);
      if (!emittedToken && options.onToken && result) {
        options.onToken(result);
      }
      return {
        text: result,
        inputTokens: Math.ceil(prompt.length / 4),
        outputTokens: Math.ceil(result.length / 4),
      };
    }

    const fallback = 'Embedded local model response: OK';
    options.onToken?.(fallback);
    return {
      text: fallback,
      inputTokens: Math.ceil(prompt.length / 4),
      outputTokens: 8,
    };
  }
}

/**
 * Format conversation and tools into a single prompt for local quantized models.
 */
export function formatLlamaPrompt(messages: ChatMessage[], tools?: ToolSpec[]): string {
  let prompt = '';

  let systemText = '';
  if (tools && tools.length > 0) {
    systemText += `You are an AI assistant in Microsoft Excel. You can call tools to read and modify the workbook.\n`;
    systemText += `Available tools (in JSON schema):\n`;
    systemText += JSON.stringify(
      tools.map((t) => ({
        name: t.name,
        description: t.description,
        parameters: t.inputSchema,
      })),
      null,
      2,
    );
    systemText += `\n\nWhen you need to call a tool, you MUST emit a block exactly like this:\n`;
    systemText += `<tool_call>\n{"name": "tool_name", "arguments": {"arg_key": "arg_value"}}\n</tool_call>\n`;
    systemText += `You may answer with normal text before or after tool calls.\n\n`;
  }

  for (const msg of messages) {
    if (msg.role === 'system') {
      systemText += `${msg.content}\n`;
    }
  }

  if (systemText.trim()) {
    prompt += `System: ${systemText.trim()}\n\n`;
  }

  for (const msg of messages) {
    if (msg.role === 'system') continue;
    if (msg.role === 'user') {
      prompt += `User: ${msg.content}\n`;
    } else if (msg.role === 'assistant') {
      prompt += `Assistant: ${msg.content}`;
      if (msg.toolCalls && msg.toolCalls.length > 0) {
        for (const tc of msg.toolCalls) {
          prompt += `\n<tool_call>\n{"name": "${tc.name}", "arguments": ${tc.argsJson}}\n</tool_call>`;
        }
      }
      prompt += '\n';
    } else if (msg.role === 'tool') {
      prompt += `Tool Result [${msg.toolCallId}]: ${msg.content}\n`;
    }
  }

  prompt += 'Assistant: ';
  return prompt;
}

function parseXmlScalarOrJson(raw: string): unknown {
  const trimmed = raw.trim();
  if (!trimmed) return '';
  if (
    (trimmed.startsWith('{') && trimmed.endsWith('}')) ||
    (trimmed.startsWith('[') && trimmed.endsWith(']')) ||
    trimmed === 'true' ||
    trimmed === 'false' ||
    trimmed === 'null' ||
    /^-?\d+(\.\d+)?$/.test(trimmed)
  ) {
    try {
      return JSON.parse(trimmed);
    } catch {
      return trimmed;
    }
  }
  return trimmed;
}

function parseSingleToolCallBody(inner: string): { name: string; argsJson: string } | null {
  const trimmed = inner.trim();
  if (!trimmed) return null;

  // 1. Standard JSON object: {"name": "...", "arguments": {...}} or {"tool": "...", "args": {...}}
  if (trimmed.startsWith('{')) {
    try {
      const parsed = JSON.parse(trimmed) as {
        name?: string;
        tool?: string;
        arguments?: unknown;
        args?: unknown;
        parameters?: unknown;
      };
      const name = parsed.name || parsed.tool;
      const args = parsed.arguments ?? parsed.args ?? parsed.parameters ?? {};
      if (name && typeof name === 'string') {
        return {
          name: name.trim(),
          argsJson: typeof args === 'string' ? args : JSON.stringify(args),
        };
      }
    } catch {
      // Fall through to XML/tag parsers
    }
  }

  // 2. <function=tool_name><parameter=key>val</parameter>...</function>
  const fnTagMatch = /<function=([a-zA-Z0-9_:-]+)>([\s\S]*?)(?:<\/function>|$)/i.exec(trimmed);
  if (fnTagMatch && fnTagMatch[1]) {
    const name = fnTagMatch[1].trim();
    const body = fnTagMatch[2] ?? '';
    const args: Record<string, unknown> = {};
    const paramRegex = /<parameter=([a-zA-Z0-9_:-]+)>([\s\S]*?)<\/parameter>/gi;
    let pMatch: RegExpExecArray | null;
    while ((pMatch = paramRegex.exec(body)) !== null) {
      if (pMatch[1]) {
        args[pMatch[1].trim()] = parseXmlScalarOrJson(pMatch[2] ?? '');
      }
    }
    return { name, argsJson: JSON.stringify(args) };
  }

  // 3. GLM-4 / Qwen / Flint XML format:
  //    e.g. "read_range<arg_key>range</arg_key><arg_value>Sheet1!A1:B5</arg_value>"
  //    or   "get_active_sheet</arg_value>"
  //    or   "get_workbook_map"
  //    or   "read_range {\"range\":\"Sheet1!A1:B5\"}"
  const leadingNameMatch = /^([a-zA-Z][a-zA-Z0-9_:-]*)\b([\s\S]*)$/.exec(trimmed);
  if (leadingNameMatch && leadingNameMatch[1]) {
    const name = leadingNameMatch[1].trim();
    const rest = (leadingNameMatch[2] ?? '').trim();

    if (rest.includes('<arg_key>')) {
      const args: Record<string, unknown> = {};
      const argPairRegex = /<arg_key>([\s\S]*?)<\/arg_key>\s*<arg_value>([\s\S]*?)(?:<\/arg_value>|$)/gi;
      let aMatch: RegExpExecArray | null;
      while ((aMatch = argPairRegex.exec(rest)) !== null) {
        const key = aMatch[1]?.trim();
        if (key) {
          args[key] = parseXmlScalarOrJson(aMatch[2] ?? '');
        }
      }
      return { name, argsJson: JSON.stringify(args) };
    }

    if (rest.startsWith('{')) {
      try {
        const parsedArgs = JSON.parse(rest);
        return { name, argsJson: JSON.stringify(parsedArgs) };
      } catch {
        // ignore
      }
    }

    // Bare tool name or trailing closing fragments like "</arg_value>"
    if (rest === '' || /^(?:<\/[a-zA-Z0-9_:-]+>\s*)+$/i.test(rest)) {
      return { name, argsJson: '{}' };
    }
  }

  return null;
}

/**
 * Extract <think>...</think> or <thinking>...</thinking> blocks (including unclosed
 * trailing blocks when generation hits max_tokens mid-thought) from raw model text.
 */
export function extractThinking(rawText: string): { cleanText: string; reasoning: string } {
  const parts: string[] = [];
  const closedRegex = /<(?:think|thinking)>([\s\S]*?)<\/(?:think|thinking)>/gi;
  let match: RegExpExecArray | null;
  while ((match = closedRegex.exec(rawText)) !== null) {
    const inner = match[1]?.trim();
    if (inner) parts.push(inner);
  }

  let cleanText = rawText.replace(closedRegex, '');

  // Handle unclosed trailing <think> or <thinking> block (e.g. truncated by max_tokens
  // or when a small local model omits </think> before emitting <tool_call>)
  const unclosedRegex = /<(?:think|thinking)>([\s\S]*?)(?=<tool_call>|$)/i;
  const unclosedMatch = unclosedRegex.exec(cleanText);
  if (unclosedMatch) {
    const inner = unclosedMatch[1]?.trim();
    if (inner) parts.push(inner);
    cleanText = cleanText.replace(unclosedRegex, '');
  }

  return {
    cleanText: cleanText.trim(),
    reasoning: parts.join('\n\n'),
  };
}

/**
 * Extract structured tool calls from model output.
 */
export function extractToolCalls(rawText: string): { cleanText: string; toolCalls: ToolCall[] } {
  const toolCalls: ToolCall[] = [];
  let cleanText = rawText;

  // 1. Check for <tool_call>...</tool_call> tags (including unclosed trailing <tool_call>)
  const tagRegex = /<tool_call>([\s\S]*?)(?:<\/tool_call>|$)/gi;
  let match: RegExpExecArray | null;
  while ((match = tagRegex.exec(rawText)) !== null) {
    const inner = match[1]?.trim();
    if (!inner) continue;
    const parsed = parseSingleToolCallBody(inner);
    if (parsed) {
      toolCalls.push({
        id: `call_${crypto.randomUUID().slice(0, 8)}`,
        name: parsed.name,
        argsJson: parsed.argsJson,
      });
    }
  }

  cleanText = cleanText.replace(tagRegex, '').trim();

  // 2. Fallback: check for markdown ```json ... ``` blocks containing {"tool": ...} or {"name": ...}
  if (toolCalls.length === 0) {
    const codeBlockRegex = /```(?:json)?\s*([\s\S]*?)\s*```/gi;
    let codeMatch: RegExpExecArray | null;
    while ((codeMatch = codeBlockRegex.exec(rawText)) !== null) {
      const block = codeMatch[1]?.trim();
      if (!block) continue;
      try {
        const parsed = JSON.parse(block) as {
          name?: string;
          tool?: string;
          arguments?: unknown;
          args?: unknown;
        };
        const name = parsed.name || parsed.tool;
        const args = parsed.arguments ?? parsed.args;
        if (name && typeof name === 'string' && args !== undefined) {
          toolCalls.push({
            id: `call_${crypto.randomUUID().slice(0, 8)}`,
            name,
            argsJson: typeof args === 'string' ? args : JSON.stringify(args),
          });
          cleanText = cleanText.replace(codeMatch[0], '').trim();
        }
      } catch {
        // Not a JSON tool block
      }
    }
  }

  // 3. Fallback: bare <function=tool_name>... or tool_name<arg_key>... emitted without outer <tool_call> wrapper
  if (toolCalls.length === 0 && (/<function=/i.test(rawText) || /<arg_key>/i.test(rawText))) {
    const parsed = parseSingleToolCallBody(rawText);
    if (parsed) {
      toolCalls.push({
        id: `call_${crypto.randomUUID().slice(0, 8)}`,
        name: parsed.name,
        argsJson: parsed.argsJson,
      });
      cleanText = '';
    }
  }

  return { cleanText, toolCalls };
}

/**
 * Provider implementation executing models directly in-process via GGUF bindings.
 */
export class EmbeddedLlamaProvider implements Provider {
  readonly id = 'embedded';
  readonly model: string;
  readonly supportsTools = true;
  private readonly resolvedPath: string;
  private readonly backend: LlamaEngineBackend;
  private readonly contextSize?: number;
  private readonly threads?: number;
  private readonly gpuLayers?: number;

  constructor(
    config: ProviderConfig,
    options?: {
      backend?: LlamaEngineBackend;
      skipPathValidation?: boolean;
    },
  ) {
    this.model = config.model;
    this.contextSize = config.contextSize;
    this.threads = config.threads;
    this.gpuLayers = config.gpuLayers;

    if (options?.skipPathValidation) {
      this.resolvedPath = config.modelPath || config.model;
    } else {
      this.resolvedPath = resolveModelPath(config.model, config.modelPath);
    }

    this.backend = options?.backend ?? new NodeLlamaCppBackend();
  }

  getResolvedPath(): string {
    return this.resolvedPath;
  }

  async chat(options: ChatRequestOptions): Promise<ChatResponse> {
    if (options.abortSignal?.aborted) {
      throw new Error('AbortError: Generation canceled by user');
    }

    const prompt = formatLlamaPrompt(options.messages, options.tools);
    const splitter = options.onDelta ? new ThinkStreamSplitter(options.onDelta) : undefined;

    const result = await this.backend.generate(prompt, {
      modelPath: this.resolvedPath,
      contextSize: this.contextSize,
      threads: this.threads,
      gpuLayers: this.gpuLayers,
      temperature: options.temperature,
      maxTokens: options.maxTokens,
      abortSignal: options.abortSignal,
      onToken: splitter ? (chunk) => splitter.pushContent(chunk) : undefined,
    });
    splitter?.flush();

    const { cleanText: withoutThinking, reasoning } = extractThinking(result.text);
    const { cleanText, toolCalls } = extractToolCalls(withoutThinking);

    const out: ChatResponse = {
      text: cleanText,
      toolCalls,
      usage: {
        inputTokens: result.inputTokens ?? Math.ceil(prompt.length / 4),
        outputTokens: result.outputTokens ?? Math.ceil(result.text.length / 4),
      },
      stopReason:
        toolCalls.length > 0
          ? 'tool_use'
          : !cleanText.trim() && reasoning.length > 0
            ? 'max_tokens'
            : 'end',
    };
    if (reasoning.length > 0) {
      out.reasoning = reasoning;
    }
    return out;
  }
}
