import { existsSync } from 'node:fs';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { describe, expect, it } from 'vitest';
import { NodeLlamaCppBackend } from '../src/providers/embedded.js';

/**
 * L1 against the real optional native binding. Opt-in, because it needs a
 * multi-gigabyte GGUF and node-llama-cpp, neither of which is in the lockfile:
 *
 *   EXCELAI_TEST_GGUF=C:\models\some-model.gguf
 *   EXCELAI_TEST_NODE_LLAMA_CPP=C:\scratch\node_modules\node-llama-cpp   (optional;
 *     defaults to resolving `node-llama-cpp` normally)
 *
 * The module is wrapped only to record the native objects the production
 * backend creates; every call goes through to the real binding.
 */
const gguf = process.env.EXCELAI_TEST_GGUF;
const runNative = !!gguf && existsSync(gguf);

async function recordingModule() {
  const dir = process.env.EXCELAI_TEST_NODE_LLAMA_CPP;
  // @ts-ignore optional native dependency, absent from the lockfile
  const real: any = dir ? await import(pathToFileURL(join(dir, 'dist', 'index.js')).href) : await import('node-llama-cpp');
  const log = { models: [] as any[], contexts: [] as any[], sessions: [] as any[], llamas: [] as any[] };

  const intercept = <T extends object>(target: T, name: string, after: (result: any) => void): T =>
    new Proxy(target, {
      get(t, p) {
        const value = Reflect.get(t, p, t);
        if (typeof value !== 'function') return value;
        if (p !== name) return value.bind(t);
        return async (...args: unknown[]) => {
          const result = await value.apply(t, args);
          after(result);
          return result;
        };
      },
    });

  class RecordingSession extends real.LlamaChatSession {
    constructor(options: unknown) {
      super(options);
      log.sessions.push(this);
    }
  }
  const mod = {
    LlamaChatSession: RecordingSession,
    getLlama: async (options: unknown) => {
      const llama = await real.getLlama(options);
      log.llamas.push(llama);
      return intercept(llama, 'loadModel', (model) => {
        const createContext = model.createContext.bind(model);
        model.createContext = async (options: unknown) => {
          const ctx = await createContext(options);
          log.contexts.push(ctx);
          return ctx;
        };
        log.models.push(model);
      });
    },
  };
  return { mod, log };
}

describe.skipIf(!runNative)('NodeLlamaCppBackend against the real node-llama-cpp binding (L1)', () => {
  const base = { contextSize: 512, gpuLayers: 0, temperature: 0 };

  it(
    'disposes contexts and sessions on completion, cancel and error; evicts and disposes models',
    async () => {
      const { mod, log } = await recordingModule();
      const backend = new NodeLlamaCppBackend({ loadModule: async () => mod, maxModels: 1 });
      const modelPath = gguf!;

      // Completion
      const done = await backend.generate('Say OK.', { ...base, modelPath, maxTokens: 4 });
      expect(typeof done.text).toBe('string');
      expect(log.contexts).toHaveLength(1);
      expect(log.contexts[0].disposed).toBe(true);
      expect(log.sessions[0].disposed).toBe(true);
      expect(log.models[0].disposed).toBe(false);

      // Cancel after the first token
      const controller = new AbortController();
      await expect(
        backend.generate('Count from 1 to 100.', {
          ...base,
          modelPath,
          maxTokens: 256,
          abortSignal: controller.signal,
          onToken: () => controller.abort(),
        }),
      ).rejects.toThrow(/AbortError/);
      expect(log.contexts).toHaveLength(2);
      expect(log.contexts[1].disposed).toBe(true);
      expect(log.sessions[1].disposed).toBe(true);

      // A native-side failure: the context cannot be allocated. (Neither a
      // throwing onToken nor an oversized prompt is an error path on the real
      // binding: it stops early or shifts the context and resolves.)
      await expect(
        backend.generate('Say OK.', { ...base, contextSize: 100_000_000, modelPath, maxTokens: 2 }),
      ).rejects.toThrow(/Embedded GGUF execution error/);
      expect(log.contexts).toHaveLength(2);
      expect(log.models).toHaveLength(1);

      // A second cache key (same file, other separators) evicts the first
      // model. It is disposed at once only if every generation above,
      // including the failed one, released it.
      const otherKey = modelPath.includes('\\') ? modelPath.replaceAll('\\', '/') : modelPath.replaceAll('/', '\\');
      await backend.generate('Say OK.', { ...base, modelPath: otherKey, maxTokens: 2 });
      await new Promise((r) => setTimeout(r, 50));
      expect(log.models).toHaveLength(2);
      expect(log.models[0].disposed).toBe(true);
      expect(backend.cachedModelPaths()).toEqual([otherKey]);

      await backend.dispose();
      expect(log.models[1].disposed).toBe(true);
      expect(log.llamas[0].disposed).toBe(true);
    },
    600_000,
  );
});
