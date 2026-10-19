import { describe, expect, it } from 'vitest';
import { assertLoopbackUrl, NonLocalConnectionBlockedError } from '../src/networkGuard.js';
import { OpenAIProvider } from '../src/providers/openai.js';
import { AnthropicProvider } from '../src/providers/anthropic.js';

describe('Loopback Network Guard', () => {
  it('allows standard loopback addresses (127.0.0.1, localhost, [::1])', () => {
    const valid = [
      'http://127.0.0.1:1234/v1',
      'http://127.0.0.1:8923/api/health',
      'http://localhost:11434/v1',
      'http://localhost:8923/taskpane.html',
      'http://[::1]:1234/v1',
      'ws://127.0.0.1:8923/ws',
      'ws://localhost:8923/ws',
      'http://127.0.0.2:8080/v1',
      'http://127.255.255.254:1234',
    ];

    for (const url of valid) {
      expect(() => assertLoopbackUrl(url)).not.toThrow();
      expect(assertLoopbackUrl(url)).toBeInstanceOf(URL);
    }
  });

  it('rejects external cloud domains with NonLocalConnectionBlockedError', () => {
    const blockedDomains = [
      'https://api.openai.com/v1',
      'https://api.anthropic.com/v1/messages',
      'https://api.groq.com/openai/v1',
      'http://example.com',
      'https://telemetry.microsoft.com',
    ];

    for (const url of blockedDomains) {
      expect(() => assertLoopbackUrl(url)).toThrow(NonLocalConnectionBlockedError);
    }
  });

  it('rejects private LAN and public IP addresses with NonLocalConnectionBlockedError', () => {
    const blockedIps = [
      'http://192.168.1.1:1234',
      'http://10.0.0.1:8080',
      'http://172.16.0.5:1234',
      'http://169.254.1.1',
      'http://8.8.8.8:1234',
      'http://1.1.1.1',
    ];

    for (const url of blockedIps) {
      expect(() => assertLoopbackUrl(url)).toThrow(NonLocalConnectionBlockedError);
    }
  });

  it('rejects malformed URLs', () => {
    expect(() => assertLoopbackUrl('not-a-valid-url')).toThrow(NonLocalConnectionBlockedError);
    expect(() => assertLoopbackUrl('')).toThrow(NonLocalConnectionBlockedError);
  });

  it('OpenAIProvider rejects non-loopback baseUrl immediately in constructor', () => {
    expect(
      () => new OpenAIProvider('test-model', 'test-key', 'https://api.openai.com/v1'),
    ).toThrow(NonLocalConnectionBlockedError);

    expect(
      () => new OpenAIProvider('test-model', 'test-key', 'http://192.168.1.50:1234/v1'),
    ).toThrow(NonLocalConnectionBlockedError);
  });

  it('OpenAIProvider accepts loopback baseUrl and defaults to LM Studio 127.0.0.1:1234', () => {
    const defaultProvider = new OpenAIProvider('local-model');
    expect(defaultProvider).toBeDefined();

    const customLocal = new OpenAIProvider('local-model', 'test', 'http://127.0.0.1:8080/v1');
    expect(customLocal).toBeDefined();
  });

  it('AnthropicProvider rejects non-loopback baseUrl in constructor', () => {
    expect(
      () => new AnthropicProvider('claude-3', 'test-key', 'https://api.anthropic.com'),
    ).toThrow(NonLocalConnectionBlockedError);

    expect(
      () => new AnthropicProvider('claude-3', 'test-key', 'http://10.0.0.2:1234'),
    ).toThrow(NonLocalConnectionBlockedError);
  });

  it('AnthropicProvider accepts loopback baseUrl', () => {
    const provider = new AnthropicProvider('local-claude-proxy', 'test', 'http://127.0.0.1:1234/v1');
    expect(provider).toBeDefined();
  });
});
