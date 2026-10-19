import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import {
  EmbeddedLlamaProvider,
  extractToolCalls,
  formatLlamaPrompt,
  getModelsDir,
  listLocalGgufModels,
  MockLlamaBackend,
  resolveModelPath,
} from '../src/providers/embedded.js';
import { ProviderError } from '../src/providers/types.js';

describe('Embedded LLaMA Provider & Local GGUF Engine', () => {
  let tempDir: string;

  beforeEach(() => {
    tempDir = join(tmpdir(), `excelai-test-${crypto.randomUUID()}`);
    mkdirSync(tempDir, { recursive: true });
  });

  afterEach(() => {
    try {
      rmSync(tempDir, { recursive: true, force: true });
    } catch {
      // ignore
    }
  });

  describe('listLocalGgufModels', () => {
    it('returns empty list for non-existent directory', async () => {
      const nonExistent = join(tempDir, 'does-not-exist');
      const models = await listLocalGgufModels(nonExistent);
      expect(models).toEqual([]);
    });

    it('scans directory and returns only .gguf files sorted', async () => {
      writeFileSync(join(tempDir, 'model-b.gguf'), 'dummy-b');
      writeFileSync(join(tempDir, 'model-a.gguf'), 'dummy-a');
      writeFileSync(join(tempDir, 'notes.txt'), 'notes');
      writeFileSync(join(tempDir, 'config.json'), '{}');

      const models = await listLocalGgufModels(tempDir);
      expect(models).toEqual(['model-a.gguf', 'model-b.gguf']);
    });
  });

  describe('resolveModelPath & Security Sandboxing', () => {
    it('resolves relative gguf file within models directory', () => {
      const dummyPath = join(tempDir, 'test-model.gguf');
      writeFileSync(dummyPath, 'gguf-content');

      // Test with custom models dir
      const oldEnv = process.env.EXCELAI_MODELS_DIR;
      process.env.EXCELAI_MODELS_DIR = tempDir;
      try {
        const resolved = resolveModelPath('test-model');
        expect(resolved.toLowerCase()).toBe(dummyPath.toLowerCase());
      } finally {
        process.env.EXCELAI_MODELS_DIR = oldEnv;
      }
    });

    it('resolves absolute path when file exists and has .gguf extension', () => {
      const dummyPath = join(tempDir, 'custom.gguf');
      writeFileSync(dummyPath, 'gguf-content');

      const resolved = resolveModelPath(dummyPath);
      expect(resolved.toLowerCase()).toBe(dummyPath.toLowerCase());
    });

    it('rejects files without .gguf extension', () => {
      const dummyPath = join(tempDir, 'dangerous.bin');
      writeFileSync(dummyPath, 'not-a-gguf');

      expect(() => resolveModelPath(dummyPath)).toThrow(ProviderError);
      expect(() => resolveModelPath(dummyPath)).toThrow(/requires a \.gguf model file/);
    });

    it('rejects non-existent model files', () => {
      const nonExistent = join(tempDir, 'missing.gguf');
      expect(() => resolveModelPath(nonExistent)).toThrow(ProviderError);
      expect(() => resolveModelPath(nonExistent)).toThrow(/Embedded model file not found/);
    });

    it('sanitizes relative paths to prevent directory traversal', () => {
      const dummyPath = join(tempDir, 'traversal.gguf');
      writeFileSync(dummyPath, 'gguf-content');

      const oldEnv = process.env.EXCELAI_MODELS_DIR;
      process.env.EXCELAI_MODELS_DIR = tempDir;
      try {
        // Passing '../../traversal.gguf' should only take basename inside models directory
        const resolved = resolveModelPath('../../traversal.gguf');
        expect(resolved.toLowerCase()).toBe(dummyPath.toLowerCase());
      } finally {
        process.env.EXCELAI_MODELS_DIR = oldEnv;
      }
    });
  });

  describe('Prompt Formatting & Tool Calling', () => {
    it('formats system prompt with tools and messages into llama format', () => {
      const prompt = formatLlamaPrompt(
        [
          { role: 'system', content: 'You are an expert.' },
          { role: 'user', content: 'Sum columns A and B.' },
          {
            role: 'assistant',
            content: 'Calculating...',
            toolCalls: [{ id: 'c1', name: 'get_used_range', argsJson: '{"sheetName": "Sheet1"}' }],
          },
          { role: 'tool', toolCallId: 'c1', content: '{"range": "A1:B10"}' },
        ],
        [
          {
            name: 'get_used_range',
            description: 'Get used range of a sheet',
            inputSchema: { type: 'object' },
          },
        ],
      );

      expect(prompt).toContain('Available tools');
      expect(prompt).toContain('get_used_range');
      expect(prompt).toContain('User: Sum columns A and B.');
      expect(prompt).toContain('<tool_call>');
      expect(prompt).toContain('Tool Result [c1]: {"range": "A1:B10"}');
      expect(prompt.endsWith('Assistant: ')).toBe(true);
    });

    it('extracts structured <tool_call> tags from output', () => {
      const raw = `I will inspect the worksheet.\n<tool_call>\n{"name": "get_active_sheet_summary", "arguments": {}}\n</tool_call>\nDone.`;
      const { cleanText, toolCalls } = extractToolCalls(raw);

      expect(toolCalls).toHaveLength(1);
      expect(toolCalls[0].name).toBe('get_active_sheet_summary');
      expect(toolCalls[0].argsJson).toBe('{}');
      expect(cleanText).toBe('I will inspect the worksheet.\n\nDone.');
    });

    it('extracts fallback markdown json code blocks for tool use', () => {
      const raw = `Calling range tool:\n\`\`\`json\n{"tool": "get_used_range", "args": {"sheet": "Sheet1"}}\n\`\`\``;
      const { cleanText, toolCalls } = extractToolCalls(raw);

      expect(toolCalls).toHaveLength(1);
      expect(toolCalls[0].name).toBe('get_used_range');
      expect(JSON.parse(toolCalls[0].argsJson)).toEqual({ sheet: 'Sheet1' });
      expect(cleanText).toBe('Calling range tool:');
    });
  });

  describe('EmbeddedLlamaProvider Execution & Cancellation', () => {
    it('executes chat turn and returns clean response', async () => {
      const dummyPath = join(tempDir, 'qwen.gguf');
      writeFileSync(dummyPath, 'mock-gguf');

      const mockBackend = new MockLlamaBackend(async (prompt) => {
        expect(prompt).toContain('User: Hello');
        return 'Hello! How can I help with your spreadsheet?';
      });

      const provider = new EmbeddedLlamaProvider(
        { kind: 'embedded', model: dummyPath },
        { backend: mockBackend },
      );

      const res = await provider.chat({
        messages: [{ role: 'user', content: 'Hello' }],
      });

      expect(res.text).toBe('Hello! How can I help with your spreadsheet?');
      expect(res.stopReason).toBe('end');
      expect(res.toolCalls).toHaveLength(0);
      expect(res.usage.outputTokens).toBeGreaterThan(0);
    });

    it('handles tool calls returned by backend', async () => {
      const dummyPath = join(tempDir, 'qwen.gguf');
      writeFileSync(dummyPath, 'mock-gguf');

      const mockBackend = new MockLlamaBackend(async () => {
        return `<tool_call>\n{"name": "get_active_sheet_summary", "arguments": {"includeData": true}}\n</tool_call>`;
      });

      const provider = new EmbeddedLlamaProvider(
        { kind: 'embedded', model: dummyPath },
        { backend: mockBackend },
      );

      const res = await provider.chat({
        messages: [{ role: 'user', content: 'Inspect the sheet' }],
      });

      expect(res.stopReason).toBe('tool_use');
      expect(res.toolCalls).toHaveLength(1);
      expect(res.toolCalls[0].name).toBe('get_active_sheet_summary');
      expect(JSON.parse(res.toolCalls[0].argsJson)).toEqual({ includeData: true });
    });

    it('immediately aborts when AbortSignal is already triggered or triggered in-flight', async () => {
      const dummyPath = join(tempDir, 'qwen.gguf');
      writeFileSync(dummyPath, 'mock-gguf');

      const mockBackend = new MockLlamaBackend(async (_prompt, signal) => {
        if (signal?.aborted) throw new Error('AbortError');
        return 'Should not reach here';
      });

      const provider = new EmbeddedLlamaProvider(
        { kind: 'embedded', model: dummyPath },
        { backend: mockBackend },
      );

      const controller = new AbortController();
      controller.abort();

      await expect(
        provider.chat({
          messages: [{ role: 'user', content: 'Hello' }],
          abortSignal: controller.signal,
        }),
      ).rejects.toThrow(/AbortError/);
    });
  });
});
