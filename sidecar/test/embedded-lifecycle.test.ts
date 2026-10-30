import { describe, expect, it } from 'vitest';
import { NodeLlamaCppBackend } from '../src/providers/embedded.js';

/**
 * A stand-in for the node-llama-cpp module that records every native object
 * the production backend creates, so disposal can be asserted on each path.
 */
function fakeLlamaModule(
  prompt: (args: { signal?: AbortSignal; onTextChunk?: (c: string) => void }) => Promise<string> = async () => 'ok',
  hooks: { createContext?: () => void } = {},
) {
  class Resource {
    disposed = false;
    dispose(): void {
      this.disposed = true;
    }
  }
  const log = {
    llamas: [] as Resource[],
    models: [] as Array<Resource & { path: string }>,
    contexts: [] as Resource[],
    sessions: [] as Resource[],
  };
  class FakeContext extends Resource {
    getSequence(): object {
      return {};
    }
  }
  class FakeModel extends Resource {
    constructor(readonly path: string) {
      super();
    }
    async createContext(): Promise<FakeContext> {
      hooks.createContext?.();
      const ctx = new FakeContext();
      log.contexts.push(ctx);
      return ctx;
    }
  }
  class FakeLlama extends Resource {
    async loadModel({ modelPath }: { modelPath: string }): Promise<FakeModel> {
      await Promise.resolve();
      const model = new FakeModel(modelPath);
      log.models.push(model);
      return model;
    }
  }
  class LlamaChatSession extends Resource {
    constructor() {
      super();
      log.sessions.push(this);
    }
    prompt(_text: string, opts: { signal?: AbortSignal; onTextChunk?: (c: string) => void }): Promise<string> {
      return prompt(opts);
    }
  }
  const mod = {
    getLlama: async () => {
      const llama = new FakeLlama();
      log.llamas.push(llama);
      return llama;
    },
    LlamaChatSession,
  };
  return { mod, log };
}

function deferred<T>() {
  let resolve!: (v: T) => void;
  const promise = new Promise<T>((r) => (resolve = r));
  return { promise, resolve };
}

describe('NodeLlamaCppBackend native lifecycle (L1)', () => {
  it('disposes the context and session after a completed generation but keeps the model warm', async () => {
    const { mod, log } = fakeLlamaModule();
    const backend = new NodeLlamaCppBackend({ loadModule: async () => mod });

    await expect(backend.generate('hi', { modelPath: 'a.gguf' })).resolves.toMatchObject({ text: 'ok' });
    await backend.generate('again', { modelPath: 'a.gguf' });

    expect(log.contexts).toHaveLength(2);
    expect(log.contexts.every((c) => c.disposed)).toBe(true);
    expect(log.sessions.every((s) => s.disposed)).toBe(true);
    expect(log.models).toHaveLength(1);
    expect(log.models[0]!.disposed).toBe(false);
  });

  it('disposes the context and session when the generation is cancelled', async () => {
    const { mod, log } = fakeLlamaModule(
      ({ signal }) =>
        new Promise((_, reject) => {
          signal?.addEventListener('abort', () => reject(Object.assign(new Error('aborted'), { name: 'AbortError' })));
        }),
    );
    const backend = new NodeLlamaCppBackend({ loadModule: async () => mod });
    const controller = new AbortController();

    const run = backend.generate('hi', { modelPath: 'a.gguf', abortSignal: controller.signal });
    await new Promise((r) => setTimeout(r, 5));
    controller.abort();

    await expect(run).rejects.toThrow(/AbortError/);
    expect(log.contexts).toHaveLength(1);
    expect(log.contexts[0]!.disposed).toBe(true);
    expect(log.sessions[0]!.disposed).toBe(true);
  });

  it('disposes the context and session when the generation fails', async () => {
    const { mod, log } = fakeLlamaModule(async () => {
      throw new Error('decode failed');
    });
    const backend = new NodeLlamaCppBackend({ loadModule: async () => mod });

    await expect(backend.generate('hi', { modelPath: 'a.gguf' })).rejects.toThrow(/decode failed/);
    expect(log.contexts[0]!.disposed).toBe(true);
    expect(log.sessions[0]!.disposed).toBe(true);
  });

  it('releases the model when context creation fails', async () => {
    let fail = true;
    const { mod, log } = fakeLlamaModule(undefined, {
      createContext: () => {
        if (fail) throw new Error('out of memory');
      },
    });
    const backend = new NodeLlamaCppBackend({ loadModule: async () => mod, maxModels: 1 });

    await expect(backend.generate('hi', { modelPath: 'a.gguf' })).rejects.toThrow(/out of memory/);
    fail = false;
    // The failed generation released its hold, so eviction disposes A at once.
    await backend.generate('hi', { modelPath: 'b.gguf' });
    await new Promise((r) => setTimeout(r, 0));
    expect(log.models.find((m) => m.path === 'a.gguf')!.disposed).toBe(true);
  });

  it('bounds the model cache with LRU eviction and disposes evicted models', async () => {
    const { mod, log } = fakeLlamaModule();
    const backend = new NodeLlamaCppBackend({ loadModule: async () => mod, maxModels: 2 });

    await backend.generate('1', { modelPath: 'a.gguf' });
    await backend.generate('2', { modelPath: 'b.gguf' });
    await backend.generate('3', { modelPath: 'a.gguf' }); // a is now most recent
    await backend.generate('4', { modelPath: 'c.gguf' }); // evicts b
    await new Promise((r) => setTimeout(r, 0));

    expect(backend.cachedModelPaths()).toEqual(['a.gguf', 'c.gguf']);
    const byPath = (p: string) => log.models.filter((m) => m.path === p);
    expect(byPath('a.gguf')).toHaveLength(1);
    expect(byPath('a.gguf')[0]!.disposed).toBe(false);
    expect(byPath('b.gguf')[0]!.disposed).toBe(true);
    expect(byPath('c.gguf')[0]!.disposed).toBe(false);
  });

  it('defers disposing an evicted model until its in-flight generation finishes', async () => {
    const gate = deferred<string>();
    let calls = 0;
    const { mod, log } = fakeLlamaModule(() => (calls++ === 0 ? gate.promise : Promise.resolve('b')));
    const backend = new NodeLlamaCppBackend({ loadModule: async () => mod, maxModels: 1 });

    const slow = backend.generate('a', { modelPath: 'a.gguf' });
    await new Promise((r) => setTimeout(r, 5));
    await backend.generate('b', { modelPath: 'b.gguf' });

    const modelA = log.models.find((m) => m.path === 'a.gguf')!;
    expect(backend.cachedModelPaths()).toEqual(['b.gguf']);
    expect(modelA.disposed).toBe(false);

    gate.resolve('a');
    await slow;
    await new Promise((r) => setTimeout(r, 0));
    expect(modelA.disposed).toBe(true);
  });

  it('loads a model once for concurrent generations on the same path', async () => {
    const { mod, log } = fakeLlamaModule();
    const backend = new NodeLlamaCppBackend({ loadModule: async () => mod });

    await Promise.all([
      backend.generate('1', { modelPath: 'a.gguf' }),
      backend.generate('2', { modelPath: 'a.gguf' }),
      backend.generate('3', { modelPath: 'a.gguf' }),
    ]);

    expect(log.llamas).toHaveLength(1);
    expect(log.models).toHaveLength(1);
  });

  it('dispose() frees every cached model and the runtime', async () => {
    const { mod, log } = fakeLlamaModule();
    const backend = new NodeLlamaCppBackend({ loadModule: async () => mod, maxModels: 2 });
    await backend.generate('1', { modelPath: 'a.gguf' });
    await backend.generate('2', { modelPath: 'b.gguf' });

    await backend.dispose();

    expect(backend.cachedModelPaths()).toEqual([]);
    expect(log.models.every((m) => m.disposed)).toBe(true);
    expect(log.llamas[0]!.disposed).toBe(true);
  });
});
