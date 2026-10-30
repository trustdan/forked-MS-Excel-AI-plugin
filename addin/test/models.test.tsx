import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { beforeEach, describe, expect, it, vi } from 'vitest';
import * as connection from '../src/connection.js';
import { ModelsScreen } from '../src/taskpane/components/settings/ModelsScreen.js';

describe('ModelsScreen Component', () => {
  beforeEach(() => {
    vi.restoreAllMocks();
  });

  it('renders provider list and allows adding an embedded GGUF model', async () => {
    const mockApi = vi.spyOn(connection, 'api').mockImplementation(async (path, init) => {
      if (path === '/api/settings/providers' && (!init || init.method === undefined)) {
        return {
          providers: [
            {
              id: 'p1',
              kind: 'openai-compatible',
              model: 'local-model',
              baseUrl: 'http://127.0.0.1:1234/v1',
            },
          ],
          activeProviderId: 'p1',
        };
      }
      if (path === '/api/providers/models' && init?.method === 'POST') {
        return { models: ['qwen2.5-7b-instruct-q4_k_m.gguf'] };
      }
      if (path === '/api/settings/providers' && init?.method === 'POST') {
        return { id: 'p2', activeProviderId: 'p1' };
      }
      return {};
    });

    const onChanged = vi.fn();
    render(<ModelsScreen onChanged={onChanged} />);

    // Wait for initial list to load
    await waitFor(() => {
      expect(screen.getByText('local-model')).toBeDefined();
      expect(screen.getByText('active')).toBeDefined();
    });

    // Click "+ Add a model"
    const addBtn = screen.getByRole('button', { name: /\+ Add a model/i });
    fireEvent.click(addBtn);

    // Select "Embedded GGUF"
    const select = screen.getByLabelText(/Provider/i) as HTMLSelectElement;
    fireEvent.change(select, { target: { value: 'embedded' } });

    // Verify hint is displayed
    expect(screen.getByText(/Runs completely in-process/i)).toBeDefined();

    // Verify "Scan local models" button appears and Base URL input is hidden
    const scanBtn = screen.getByRole('button', { name: /Scan local models/i });
    expect(scanBtn).toBeDefined();
    expect(screen.queryByLabelText(/Base URL/i)).toBeNull();

    // Click "Scan local models"
    fireEvent.click(scanBtn);

    await waitFor(() => {
      expect(mockApi).toHaveBeenCalledWith(
        '/api/providers/models',
        expect.objectContaining({ method: 'POST' }),
      );
      expect(screen.getByText(/1 models available/i)).toBeDefined();
    });
  });

  it('labels cloud-named kinds as local protocol adapters with an editable loopback URL', async () => {
    const mockApi = vi.spyOn(connection, 'api').mockImplementation(async (path, init) => {
      if (path === '/api/settings/providers' && !init?.method) {
        return { providers: [], activeProviderId: null };
      }
      if (path === '/api/settings/providers' && init?.method === 'POST') {
        return { id: 'p1', activeProviderId: 'p1' };
      }
      return {};
    });

    render(<ModelsScreen onChanged={vi.fn()} />);
    const select = (await screen.findByLabelText(/Provider/i)) as HTMLSelectElement;
    const labels = Array.from(select.options).map((o) => o.text);
    expect(labels).toContain('Anthropic Messages protocol (local server)');
    expect(labels).toContain('OpenAI API protocol (local server)');
    expect(labels).not.toContain('Anthropic');
    expect(labels).not.toContain('OpenAI');

    fireEvent.change(select, { target: { value: 'anthropic' } });
    expect(screen.getByText(/Cloud endpoints are blocked/i)).toBeDefined();
    const baseUrl = screen.getByLabelText(/Base URL/i) as HTMLInputElement;
    fireEvent.change(baseUrl, { target: { value: 'http://127.0.0.1:4000' } });
    fireEvent.change(screen.getByLabelText(/^Model/i), { target: { value: 'claude-local' } });
    fireEvent.click(screen.getByRole('button', { name: /Add model/i }));

    await waitFor(() => {
      const post = mockApi.mock.calls.find(
        ([p, i]) => p === '/api/settings/providers' && i?.method === 'POST',
      );
      expect(JSON.parse(String(post?.[1]?.body))).toMatchObject({
        kind: 'anthropic',
        baseUrl: 'http://127.0.0.1:4000',
      });
    });
  });

  it('shows a settings-file problem and dismisses it', async () => {
    let problem: object | undefined = {
      kind: 'unreadable',
      detail: 'not valid JSON',
      preservedAs: 'C:/Users/me/.excelai/settings.json.corrupt-2026-10-30T12-00-00-000Z',
      at: '2026-10-30T12:00:00.000Z',
    };
    const mockApi = vi.spyOn(connection, 'api').mockImplementation(async (path, init) => {
      if (path === '/api/settings/providers' && !init?.method) {
        return { providers: [], activeProviderId: null, ...(problem ? { settingsProblem: problem } : {}) };
      }
      if (path === '/api/settings/problem' && init?.method === 'DELETE') {
        problem = undefined;
        return { ok: true };
      }
      return {};
    });

    render(<ModelsScreen onChanged={vi.fn()} />);
    const alert = await screen.findByRole('alert');
    expect(alert.textContent).toContain('unreadable');
    expect(alert.textContent).toContain('settings.json.corrupt-2026-10-30T12-00-00-000Z');

    fireEvent.click(screen.getByRole('button', { name: /Dismiss/i }));
    await waitFor(() => expect(screen.queryByRole('alert')).toBeNull());
    expect(mockApi).toHaveBeenCalledWith('/api/settings/problem', { method: 'DELETE' });
  });
});
