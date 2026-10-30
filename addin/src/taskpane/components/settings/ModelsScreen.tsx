import { useCallback, useEffect, useState } from 'react';
import { api } from '../../../connection.js';
import {
  KIND_LABELS,
  type ProviderKind,
  type ProviderList,
} from '../../types.js';

interface AddForm {
  kind: ProviderKind;
  model: string;
  apiKey: string;
  baseUrl: string;
}

const EMPTY_FORM: AddForm = {
  kind: 'openai-compatible',
  model: 'local-model',
  apiKey: 'lm-studio',
  baseUrl: 'http://127.0.0.1:1234/v1',
};

export interface ModelsScreenProps {
  onChanged: () => void;
}

export function ModelsScreen({ onChanged }: ModelsScreenProps): JSX.Element {
  const [list, setList] = useState<ProviderList>({ providers: [], activeProviderId: null });
  const [form, setForm] = useState<AddForm>(EMPTY_FORM);
  const [showAdd, setShowAdd] = useState(false);
  const [message, setMessage] = useState('');
  const [rowMessage, setRowMessage] = useState<Record<string, string>>({});
  const [suggestions, setSuggestions] = useState<string[]>([]);
  const [busyAction, setBusyAction] = useState('');

  const load = useCallback((): void => {
    void api<ProviderList>('/api/settings/providers').then(
      (l: ProviderList) => {
        setList(l);
        if (l.providers.length === 0) setShowAdd(true);
      },
      (e: unknown) => setMessage(String(e instanceof Error ? e.message : e)),
    );
  }, []);

  useEffect(load, [load]);

  const formBody = (): Record<string, unknown> => {
    const body: Record<string, unknown> = { kind: form.kind, model: form.model.trim() };
    if (form.apiKey.trim()) body.apiKey = form.apiKey.trim();
    if (form.baseUrl.trim()) body.baseUrl = form.baseUrl.trim();
    return body;
  };

  const addModel = (): void => {
    setBusyAction('add');
    void api('/api/settings/providers', { method: 'POST', body: JSON.stringify(formBody()) })
      .then(() => {
        setForm(EMPTY_FORM);
        setShowAdd(false);
        setSuggestions([]);
        setMessage('Model added.');
        load();
        onChanged();
      })
      .catch((e: unknown) => setMessage(`✗ ${e instanceof Error ? e.message : String(e)}`))
      .finally(() => setBusyAction(''));
  };

  const testCandidate = (): void => {
    setBusyAction('test');
    setMessage('Testing…');
    void api<{ ok: boolean; reply?: string; error?: string }>('/api/settings/providers/test', {
      method: 'POST',
      body: JSON.stringify(formBody()),
    })
      .then((r: { ok: boolean; reply?: string; error?: string }) =>
        setMessage(r.ok ? `✓ Connected — model replied: "${r.reply ?? ''}"` : `✗ ${r.error ?? 'Failed'}`),
      )
      .catch((e: unknown) => setMessage(`✗ ${e instanceof Error ? e.message : String(e)}`))
      .finally(() => setBusyAction(''));
  };

  const loadModels = (): void => {
    setBusyAction('models');
    setMessage('Fetching available models…');
    const body: Record<string, unknown> = { kind: form.kind };
    if (form.apiKey.trim()) body.apiKey = form.apiKey.trim();
    if (form.baseUrl.trim()) body.baseUrl = form.baseUrl.trim();
    void api<{ models: string[]; error?: string }>('/api/providers/models', {
      method: 'POST',
      body: JSON.stringify(body),
    })
      .then((r: { models: string[]; error?: string }) => {
        setSuggestions(r.models);
        setMessage(
          r.models.length > 0
            ? `${r.models.length} models available — pick one from the Model field.`
            : `✗ ${r.error ?? 'No models found'}`,
        );
      })
      .catch((e: unknown) => setMessage(`✗ ${e instanceof Error ? e.message : String(e)}`))
      .finally(() => setBusyAction(''));
  };

  const activate = (id: string): void => {
    void api(`/api/settings/providers/${id}/activate`, { method: 'POST' })
      .then(() => {
        load();
        onChanged();
      })
      .catch((e: unknown) => setMessage(`✗ ${e instanceof Error ? e.message : String(e)}`));
  };

  const remove = (id: string): void => {
    void api(`/api/settings/providers/${id}`, { method: 'DELETE' })
      .then(() => {
        load();
        onChanged();
      })
      .catch((e: unknown) => setMessage(`✗ ${e instanceof Error ? e.message : String(e)}`));
  };

  const testSaved = (id: string): void => {
    setRowMessage((m) => ({ ...m, [id]: 'Testing…' }));
    void api<{ ok: boolean; reply?: string; error?: string }>(
      `/api/settings/providers/${id}/test`,
      { method: 'POST' },
    )
      .then((r: { ok: boolean; reply?: string; error?: string }) =>
        setRowMessage((m) => ({
          ...m,
          [id]: r.ok ? `✓ "${r.reply ?? ''}"` : `✗ ${r.error ?? 'Failed'}`,
        })),
      )
      .catch((e: unknown) =>
        setRowMessage((m) => ({ ...m, [id]: `✗ ${e instanceof Error ? e.message : String(e)}` })),
      );
  };

  const dismissProblem = (): void => {
    void api('/api/settings/problem', { method: 'DELETE' })
      .then(load)
      .catch((e: unknown) => setMessage(`✗ ${e instanceof Error ? e.message : String(e)}`));
  };

  const problem = list.settingsProblem;

  return (
    <main className="settings">
      {problem && (
        <div className="banner" role="alert">
          Your settings file was {problem.kind === 'invalid' ? 'not in the expected format' : 'unreadable'} (
          {problem.detail}), so the sidecar started with empty settings. The original was kept at{' '}
          <code>{problem.preservedAs}</code>.{' '}
          <button className="ghost" onClick={dismissProblem}>
            Dismiss
          </button>
        </div>
      )}
      <h3>Your models</h3>
      {list.providers.length === 0 && <p className="hint">No models yet — add one below.</p>}
      {list.providers.map((p) => {
        const active = p.id === list.activeProviderId;
        return (
          <div key={p.id} className={`model-card${active ? ' active' : ''}`}>
            <div className="model-card-main">
              <strong>{p.model}</strong> {active && <span className="badge applied">active</span>}
              <div className="hint">
                {KIND_LABELS[p.kind]}
                {p.baseUrl ? ` · ${p.baseUrl}` : ''}
                {p.apiKey ? ` · key ${p.apiKey}` : ' · no key'}
              </div>
              {rowMessage[p.id] && <div className="hint">{rowMessage[p.id]}</div>}
            </div>
            <div className="actions">
              {!active && <button onClick={() => activate(p.id)}>Use</button>}
              <button className="ghost" onClick={() => testSaved(p.id)}>
                Test
              </button>
              <button className="ghost danger" onClick={() => remove(p.id)}>
                Delete
              </button>
            </div>
          </div>
        );
      })}

      {!showAdd ? (
        <button className="ghost" onClick={() => setShowAdd(true)}>
          + Add a model
        </button>
      ) : (
        <div className="add-form">
          <h3>Add a model</h3>
          <label>
            Provider
            <select
              value={form.kind}
              onChange={(e) => {
                setSuggestions([]);
                const nextKind = e.target.value as ProviderKind;
                setForm({
                  ...form,
                  kind: nextKind,
                  model: nextKind === 'embedded' ? (form.model || '') : form.model,
                });
              }}
            >
              {(['openai-compatible', 'embedded', 'openai', 'anthropic'] as const).map((k) => (
                <option key={k} value={k}>
                  {KIND_LABELS[k]}
                </option>
              ))}
            </select>
          </label>
          {form.kind === 'embedded' && (
            <div className="hint" style={{ marginBottom: '10px' }}>
              💡 Runs completely in-process. Place your <code>.gguf</code> model files in <code>~/.excelai/models/</code> and click <em>Scan local models</em>.
            </div>
          )}
          {(form.kind === 'openai' || form.kind === 'anthropic') && (
            <div className="hint" style={{ marginBottom: '10px' }}>
              Speaks the {form.kind === 'openai' ? 'OpenAI' : 'Anthropic Messages'} API to a server on this
              machine, such as a local gateway. Cloud endpoints are blocked.
            </div>
          )}
          {form.kind !== 'embedded' && (
            <label>
              Base URL
              <input
                value={form.baseUrl}
                placeholder="http://127.0.0.1:1234/v1"
                onChange={(e) => setForm({ ...form, baseUrl: e.target.value })}
              />
            </label>
          )}
          {form.kind === 'openai-compatible' && (
            <>
              <div style={{ display: 'flex', gap: '8px', marginBottom: '8px', flexWrap: 'wrap' }}>
                <button
                  type="button"
                  className="ghost"
                  onClick={() =>
                    setForm({
                      ...form,
                      baseUrl: 'http://127.0.0.1:1234/v1',
                      model: form.model || 'local-model',
                      apiKey: form.apiKey || 'lm-studio',
                    })
                  }
                >
                  Use LM Studio defaults
                </button>
                <button
                  type="button"
                  className="ghost"
                  onClick={() =>
                    setForm({ ...form, baseUrl: 'http://localhost:11434/v1', model: form.model || 'llama3.1' })
                  }
                >
                  Use Ollama defaults
                </button>
              </div>
            </>
          )}
          {form.kind !== 'embedded' && (
            <label>
              API key {form.kind === 'openai-compatible' && <em>(optional)</em>}
              <input
                type="password"
                value={form.apiKey}
                placeholder="stored locally by the sidecar, never sent elsewhere"
                onChange={(e) => setForm({ ...form, apiKey: e.target.value })}
              />
            </label>
          )}
          <label>
            Model {form.kind === 'embedded' ? '(.gguf filename)' : ''}
            <input
              value={form.model}
              list="model-suggestions"
              placeholder={
                form.kind === 'embedded'
                  ? 'qwen2.5-7b-instruct-q4_k_m.gguf'
                  : form.kind === 'openai-compatible'
                    ? 'llama3.1'
                    : 'model id'
              }
              onChange={(e) => setForm({ ...form, model: e.target.value })}
            />
            <datalist id="model-suggestions">
              {suggestions.map((m) => (
                <option key={m} value={m} />
              ))}
            </datalist>
          </label>
          <div className="actions">
            <button
              className="ghost"
              onClick={loadModels}
              disabled={
                busyAction !== '' ||
                (form.kind !== 'embedded' && !form.baseUrl)
              }
            >
              {busyAction === 'models'
                ? 'Fetching…'
                : form.kind === 'embedded'
                  ? 'Scan local models'
                  : 'List available models'}
            </button>
            <button className="ghost" onClick={testCandidate} disabled={!form.model || busyAction !== ''}>
              {busyAction === 'test' ? 'Testing…' : 'Test'}
            </button>
            <button onClick={addModel} disabled={!form.model || busyAction !== ''}>
              {busyAction === 'add' ? 'Adding…' : 'Add model'}
            </button>
            {list.providers.length > 0 && (
              <button className="ghost" onClick={() => setShowAdd(false)}>
                Cancel
              </button>
            )}
          </div>
        </div>
      )}
      {message && <p className="hint">{message}</p>}
    </main>
  );
}
