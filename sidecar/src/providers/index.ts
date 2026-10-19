import { AnthropicProvider } from './anthropic.js';
import { OpenAIProvider } from './openai.js';
import { ProviderError, type Provider, type ProviderConfig } from './types.js';

export * from './types.js';
export { AnthropicProvider } from './anthropic.js';
export { OpenAIProvider } from './openai.js';
export { ScriptedProvider, type ScriptedTurn } from './scripted.js';

export function createProvider(config: ProviderConfig): Provider {
  const baseUrl = config.baseUrl || 'http://127.0.0.1:1234/v1';
  const apiKey = config.apiKey || 'lm-studio';
  switch (config.kind) {
    case 'anthropic': {
      return new AnthropicProvider(config.model, apiKey, baseUrl);
    }
    case 'openai': {
      return new OpenAIProvider(config.model, apiKey, baseUrl);
    }
    case 'openai-compatible': {
      return new OpenAIProvider(config.model, apiKey, baseUrl, 'openai-compatible');
    }
  }
}

/** Retry with exponential backoff on retryable provider errors. */
export async function withRetry<T>(
  fn: () => Promise<T>,
  attempts = 3,
  baseDelayMs = 500,
): Promise<T> {
  let lastError: unknown;
  for (let i = 0; i < attempts; i++) {
    try {
      return await fn();
    } catch (e) {
      lastError = e;
      if (e instanceof Error && (e.name === 'AbortError' || e.message.toLowerCase().includes('abort'))) {
        throw e;
      }
      const retryable = e instanceof ProviderError ? e.retryable : false;
      if (!retryable || i === attempts - 1) throw e;
      await new Promise((r) => setTimeout(r, baseDelayMs * 2 ** i));
    }
  }
  throw lastError;
}
